import json
import os.path
import time

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/presentations"]

# The ID of a sample presentation.
PRESENTATION_ID = "1EpF5x79weje2_LP3hOPS4y_JCi0kxVXTc2J7Qb6SLew"


colors = [
    "e6b8af",
    "f4cccc",
    "fce5cd",
    "fff2cc",
    "d9ead3",
    "d0e0e3",
    "c9daf8",
    "cfe2f3",
    "d9d2e9",
    "ead1dc",
]


def hex_to_rgb(hex):
    return {
        key: int(hex[i : i + 2], 16) / 255
        for i, key in zip([0, 2, 4], ["red", "green", "blue"])
    }


def find_matching_element(elements, text):
    matches = [x for x in elements if text in json.dumps(x)]
    if len(matches) > 1:
        print(f'Warning: Found {len(matches)} matches for "{text}"')
    return matches[0]


def duplicate_slide(service, object_id):
    print(f"Duplicating {object_id}")
    time.sleep(0.2)
    response = (
        service.presentations()
        .batchUpdate(
            presentationId=PRESENTATION_ID,
            body={
                "requests": [
                    {
                        "duplicateObject": {
                            "objectId": object_id,
                        }
                    }
                ]
            },
        )
        .execute()
    )
    return response["replies"][0]["duplicateObject"]["objectId"]


def replace_text(service, object_ids, to_replace, replace_with):
    print(f"Replacing '{to_replace}' with '{replace_with}' in {object_ids}")
    service.presentations().batchUpdate(
        presentationId=PRESENTATION_ID,
        body={
            "requests": [
                {
                    "replaceAllText": {
                        "replaceText": replace_with,
                        "pageObjectIds": object_ids,
                        "containsText": {"text": to_replace, "matchCase": False},
                    }
                }
            ]
        },
    ).execute()


def move_slides_to_end(service, object_ids):
    print(f"Moving {object_ids} to end")
    total_num_slides = len(
        service.presentations().get(presentationId=PRESENTATION_ID).execute()["slides"]
    )
    service.presentations().batchUpdate(
        presentationId=PRESENTATION_ID,
        body={
            "requests": [
                {
                    "updateSlidesPosition": {
                        "slideObjectIds": object_ids,
                        "insertionIndex": total_num_slides,
                    }
                }
            ]
        },
    ).execute()


def is_text_box_containing(elt, text):
    has_shape_keys = "shape" in elt and "shapeType" in elt["shape"]
    is_text_box = has_shape_keys and elt["shape"]["shapeType"] == "TEXT_BOX"
    has_text_key = "text" in elt["shape"]

    if is_text_box and not has_text_key:
        print(
            "Found text box without the 'text' key. "
            "This is ok if there are weird text boxes with no text in them, "
            "and those parts of the slide will just be skipped. Elt was: "
            f"{elt}"
        )
        return False

    return is_text_box and has_text_key and text in json.dumps(elt["shape"]["text"])


def modify_background_color_of_shapes_containing(service, match_string, color_index):
    all_slides = (
        service.presentations().get(presentationId=PRESENTATION_ID).execute()["slides"]
    )
    page_elements = [y for x in all_slides for y in x["pageElements"]]
    shapes = [x for x in page_elements if is_text_box_containing(x, match_string)]
    if not shapes:
        print(
                "Warning: Found no text boxes to modify the background color in. "
                "Either there were no text boxes or else no text boxes contained the match string "
                f"'{match_string}'.\n\nPage elements were:\n\n{page_elements}")
        return

    service.presentations().batchUpdate(
        presentationId=PRESENTATION_ID,
        body={
            "requests": [
                {
                    "updateShapeProperties": {
                        "objectId": shape["objectId"],
                        "fields": "shapeBackgroundFill.solidFill.color",
                        "shapeProperties": {
                            "shapeBackgroundFill": {
                                "solidFill": {
                                    "color": {
                                        "rgbColor": hex_to_rgb(colors[color_index]),
                                    }
                                }
                            }
                        },
                    }
                }
                for shape in shapes
            ]
        },
    ).execute()


def main():
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    try:
        service = build("slides", "v1", credentials=creds)

        # Call the Slides API
        presentation = (
            service.presentations().get(presentationId=PRESENTATION_ID).execute()
        )
        slides = presentation.get("slides")

        for i in [2, 3, 4, 5, 6, 7, 8]:
            new_object_ids = [duplicate_slide(service, x["objectId"]) for x in slides]
            move_slides_to_end(service, new_object_ids)
            replace_text(service, new_object_ids, "Group 1", f"Group {i}")
            modify_background_color_of_shapes_containing(
                service,
                f"Group {i}",
                color_index=(7 * i) % len(colors),
            )
            time.sleep(5)

        return service, slides
    except HttpError as err:
        print(err)


if __name__ == "__main__":
    service, slides = main()
