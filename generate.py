from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from datetime import datetime
import os
import argparse

file_path = "/Users/bhavya/Documents/CoOp/"


def download_document_as_pdf(drive_service, document_id, fname):
    try:
        # Set the export MIME type to PDF
        mime_type = "application/pdf"

        # Export the document as a PDF
        export_request = drive_service.files().export(
            fileId=document_id, mimeType=mime_type
        )
        export_response = export_request.execute()

        # Save the exported PDF to local storage
        with open(file_path + fname, "wb+") as f:
            f.write(export_response)

        print(f"PDF saved to {file_path}")

    except HttpError as error:
        print(f"An error occurred: {error}")


def clear_document(drive_service, document_id):
    try:
        # Export the document as a PDF
        export_request = drive_service.files().delete(fileId=document_id)
        export_request.execute()
        print("Document cleared")
    except HttpError as error:
        print(f"An error occurred: {error}")


def generate_letter(service, document, fields: dict):
    main_request = {"requests": []}
    for key, value in fields.items():
        replace_request = {
            "replaceAllText": {  # Replaces all instances of text matching a criteria with replace text. # Replaces all instances of the specified text.
                "containsText": {  # A criteria that matches a specific string of text in the document. # Finds text in the document matching this substring.# Indicates whether the search should respect case: - `True`: the search is case sensitive. - `False`: the search is case insensitive.
                    "text": "{{"
                    + key
                    + "}}",  # The text to search for in the document.
                },
                "replaceText": f"{value}",  # The text that will replace the matched text.
            },
        }
        main_request["requests"].append(replace_request)
    response = (
        service.documents()
        .batchUpdate(documentId=document, body=main_request)
        .execute()
    )
    return response["documentId"]


# duplicate document in google docs with a different title
def duplicate_document(service, document_id, title):
    body = {"name": title}
    try:
        request = service.files().copy(fileId=document_id, body=body)
        response = request.execute()
        return response.get("id")
    except HttpError as err:
        print(err)


def format_qualities(qs):
    first = (
        ", ".join(qs.split("' '")).replace("'", "").replace(")", "").replace("(", "")
    )

    return first[::-1].replace(",", "dna ", 1)[::-1]


parser = argparse.ArgumentParser()
parser.add_argument("-c", "--company", help="company name")
parser.add_argument("-p", "--position", help="position")
parser.add_argument("-q", "--qualities", help="qualities")
parser.add_argument("-t", "--tech", help="tech")
parser.add_argument("-f", "--field", help="field")

args = parser.parse_args()
fields = {
    "DATE": datetime.now().strftime("%d %B, %Y"),
    "COMPANY": args.company,
    "POSITION": args.position,
    "QUALITY": format_qualities(args.qualities),
    "TECH": args.tech,
    "FIELD": args.field,
}


try:
    auth = service_account.Credentials.from_service_account_file(
        os.path.dirname(os.path.realpath(__file__)) + "/svc_acc.json"
    )
except Exception as e:
    print(f"Authentication error: {e}")
    exit(1)

try:
    service = build("docs", "v1", credentials=auth)
    drive_service = build("drive", "v3", credentials=auth)

    # Retrieve the documents contents from the Docs service.
    document = service.documents().get(documentId=os.environ["MASTER_DOC_ID"]).execute()
    title = document.get("title")
    # duplicate document
    new_doc_id = duplicate_document(
        drive_service,
        os.environ["MASTER_DOC_ID"],
        title + datetime.now().strftime("%Y-%m-%d"),
    )
    # download document as pdf
    generate_letter(service, new_doc_id, fields)
    download_document_as_pdf(
        drive_service,
        new_doc_id,
        fields["COMPANY"].replace(" ", "_").lower() + "_Bhavya_Muni_cover_letter.pdf",
    )
    clear_document(drive_service, new_doc_id)

except HttpError as err:
    print(err)
