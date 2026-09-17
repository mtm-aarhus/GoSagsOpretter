"""Koden opretter en sag i go og redigerer de forskellige mulige indstillinger ud fra køelement, der oprettes fra power automate"""

from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from OpenOrchestrator.database.queues import QueueElement
from office365.runtime.auth.client_credential import ClientCredential
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.fields.url_value import FieldUrlValue
import json
import requests
from requests_ntlm import HttpNtlmAuth
import time

def create_ntlm_session(username: str, password: str) -> requests.Session:
    session = requests.Session()
    session.auth = HttpNtlmAuth(username, password)
    return session


def create_go_case(go_api_url: str, session: requests.Session, title: str, kle: str, facet: str, samlesag: str) -> dict:
    """Opretter en ny sag i GO og returnerer response som dict."""
    url = f"{go_api_url}/geosager/_goapi/Cases"

    metadata_xml = (
        '<z:row xmlns:z="#RowsetSchema"'
        f' ows_Title="{title}"'
        ' ows_EksterntSagsID="TestSagID"'
        ' ows_EksterntSystemID="TestSystemID"'
        ' />'
    )

    payload = json.dumps({
        "CaseTypePrefix": "GEO",         # FILLER: verify correct prefix for this case type
        "MetadataXml": metadata_xml,
        "ReturnWhenCaseFullyCreated": True
    })

    headers = {"Content-Type": "application/json"}
    response = session.post(url, headers=headers, data=payload)
    # print(response.text)
    response.raise_for_status()
    return response.json()


def update_sharepoint_list_item(orchestrator_connection: OrchestratorConnection, site_url: str, list_name: str, item_id: int, fields: dict) -> None:
    """Opdaterer et eksisterende element i en SharePoint-liste."""
    certification = orchestrator_connection.get_credential("SharePointCert")
    api = orchestrator_connection.get_credential("SharePointAPI")

    cert_credentials = {
        "tenant": api.username,
        "client_id": api.password,
        "thumbprint": certification.username,
        "cert_path": certification.password
    }

    ctx = ClientContext(site_url).with_client_certificate(**cert_credentials)
    item = ctx.web.lists.get_by_title(list_name).get_item_by_id(item_id)
    for key, value in fields.items():
        item.set_property(key, value)
    item.update()
    ctx.execute_query()

def delete_case_go(go_api_url, session, sagsnummer):
    '''
    Deletes case in go
    '''
    url = f"{go_api_url}/geosager/_goapi/Cases/{sagsnummer}"
    response = session.delete(url, data= {"Data": ""}, timeout=1200)
    response.raise_for_status()
    return response.json()

def get_site_digest(site_url: str, session: requests.Session) -> str:
    """Henter et FormDigestValue for det angivne web/scope."""
    endpoint = f"{site_url}/_api/contextinfo"
    r = session.post(endpoint, headers={"Accept": "application/json; odata=verbose"})
    r.raise_for_status()
    digest = r.json()["d"]["GetContextWebInformation"]["FormDigestValue"]
    return digest

def get_list_and_id(api_url, geoid, session, geonr):
    """Henter liste og itemid til opdatering af bruger i go."""
    endpoint = f"{api_url}/cases/{geonr}/{geoid}/_goapi/Administration/ModernConfiguration"
    
    payload = {
        "providerTypes": ["ModernCase", "MoveDocument", "Insight", "SearchSystem", "UserSettings"]
    }
    
    headers = {
        "Accept": "application/json",
        "Content-Type": "application/json"
    }
    
    for attempt in range(5):
        r = session.post(endpoint, headers=headers, json=payload)
        if r.status_code == 200 and r.text.strip():
            break
        print(f"Tom response - venter 5 sekunder før forsøg {attempt+2}")
        time.sleep(5)
    else:
        raise Exception(f"GO svarede ikke på ModernConfiguration efter 5 forsøg for {geoid}")
    data = r.json()
    caselist = data.get("ModernCase").get("ItemServerUrl").split('/')[-2]
    itemid = data.get("ModernCase").get("ListItemID")
    return(caselist, itemid)

def update_case_field(api_url: str, session: requests.Session, digest: str, form_values: list, listnumber, item_id):
    """Opdaterer felt(er) i sagslisten."""
    endpoint = (
        f"{api_url}/geosager/_api/web/GetList(@a1)/items(@a2)/ValidateUpdateListItem()"
        f"?@a1='%2Fgeosager%2FLists%2F{listnumber}'&@a2='{item_id}'"
    )

    headers = {
        "Accept": "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
        "X-RequestDigest": digest,
        "X-Sp-Requestresources": f"listUrl=%2Fgeosager%2FLists%2F{listnumber}"
    }

    payload = {
        "formValues": form_values,
        "bNewDocumentUpdate": False,
        "checkInComment": None
    }
    try:
        r = session.post(endpoint, headers=headers, data=json.dumps(payload))
        r.raise_for_status()
        return True
    except:
        return False

def update_case_owner(api_url, session, case_id: str, geonr, samlesag ):

    geo_digest = get_site_digest(f"{api_url}/geosager", session)
    
    listnumber, item_id = get_list_and_id(api_url, case_id, session, geonr)

    form_values = [
    {
        "FieldName": "Sagsprofil_GEO",
        "FieldValue": "MTM CR - Affaldstilsyn|f7ec6e9c-35f8-455e-b11f-5cfa1b74ab85;",
        "HasException": False,
        "ErrorMessage": None
    },
    {
        "FieldName": "CaseCategory",
        "FieldValue": "Åben for alle",
        "HasException": False,
        "ErrorMessage": None
    },
    {
        "FieldName": "Afdeling",
        "FieldValue": "Cirkulære Ressourcer|cebf0cf1-0bf5-45f5-bfc3-19a0d6533a74;",
        "HasException": False,
        "ErrorMessage": None
    },
    {
        "FieldName": "KLENummer",
        "FieldValue": "07.17.01 Fysisk affaldstilsyn|b9811ab4-f29c-11ef-809f-0050c2490048;",
        "HasException": False,
        "ErrorMessage": None
    },
    {
        "FieldName": "Facet",
        "FieldValue": "K08 Tilsyn og håndhævelse, overholdelse af regler|47e06970-fe91-479d-86ad-0812ebc70f1b;",
        "HasException": False,
        "ErrorMessage": None
    },
    {
        "FieldName": "CustomMasterCase",
        "FieldValue": samlesag,
        "HasException": False,
        "ErrorMessage": None
    }
]

    # Opdater felt
    result = update_case_field(api_url, session, geo_digest, form_values, listnumber, item_id)
    return result

# pylint: disable-next=unused-argument
def process(orchestrator_connection: OrchestratorConnection, queue_element: QueueElement | None = None) -> None:
    """Do the primary process of the robot."""
    orchestrator_connection.log_trace("Running process.")

    queue_json = json.loads(queue_element.data)
    listId = queue_json["Id"]
    base_title = queue_json["Title"].strip()
    kle_raw = queue_json["KLE"]  #hardcodet til vi tilføjer andre afdelinger - samme for facet    
    facet_raw = queue_json["Facet"]   
    p_nummer = queue_json.get("P-nummer", "")
    adresse = queue_json.get("Adresse", "")
    postnr_by_raw = queue_json.get("PostnrBy", "")
    samlesag = queue_json.get("Samlesag", "")

    # PostnrBy arrives as a JSON string
    postnr_by_value = ""
    if postnr_by_raw:
        try:
            postnr_by_obj = json.loads(postnr_by_raw)
            postnr_by_value = postnr_by_obj.get("Value", "")
        except (json.JSONDecodeError, TypeError):
            postnr_by_value = postnr_by_raw

    # Title format: "Fysik affaldstilsyn - base_title (P-nummer) Adresse - 8000 Århus C"
    p_nummer_part = f" ({p_nummer})" if p_nummer else ""
    title = f"Affaldstilsyn - {base_title}{p_nummer_part} {adresse} - {postnr_by_value}".strip()

    go_credentials = orchestrator_connection.get_credential("GOAktApiUser") 
    go_username = go_credentials.username
    go_password = go_credentials.password
    go_api_url = orchestrator_connection.get_constant("GOApiURL").value     
    session = create_ntlm_session(go_username, go_password)

    orchestrator_connection.log_info(f"Opretter GO-sag: {title}")
    #Unused deleter, used for testing
    # delete_case_go(go_api_url= go_api_url, session= session, sagsnummer=  )

    result = create_go_case(go_api_url, session, title, kle_raw, facet_raw, samlesag)

    case_id = result.get("CaseID")
    RelativeSagsUrl = result['CaseRelativeUrl']
    case_url = f'{go_api_url}/{RelativeSagsUrl}'.strip('ad.')
    geonr = RelativeSagsUrl.split('/')[-2]
    orchestrator_connection.log_info(f"GO-sag oprettet: {case_id}")
    update = update_case_owner(session= session, api_url= go_api_url, case_id= case_id, geonr= geonr, samlesag= samlesag)
    # print(case_url)

    sp_site_url = f'{orchestrator_connection.get_constant("AarhusKommuneSharePoint").value}/Teams/tea-teamsite12556'  
    update_sharepoint_list_item(
        orchestrator_connection=orchestrator_connection,
        site_url=sp_site_url,
        list_name="Affaldstilsyn",
        item_id=int(listId),
        fields={
        "GoSag": FieldUrlValue(case_url.replace('ad.', '')),   # Description defaulter til url'en
    }
    )