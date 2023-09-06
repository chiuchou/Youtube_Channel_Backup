import os
import pickle
from google_auth_oauthlib.flow import InstalledAppFlow
import googleapiclient.discovery
import googleapiclient.errors
from openpyxl import Workbook

# Set the proxy server details
proxy_server = "http://127.0.0.1:7890"

# Set the proxy server details for both HTTP and HTTPS
os.environ["HTTP_PROXY"] = proxy_server
os.environ["HTTPS_PROXY"] = proxy_server

scopes = ["https://www.googleapis.com/auth/youtube.force-ssl"]
credentials_file = "credentials.pickle"
output_file = "video_data.xlsx"

def authenticate():
    if os.path.exists(credentials_file):
        with open(credentials_file, "rb") as f:
            credentials = pickle.load(f)
    else:
        os.environ["OAUTHLIB_INSECURE_TRANSPORT"] = "1"
        api_service_name = "youtube"
        api_version = "v3"
        client_secrets_file = "client_secret_671541520290-hofccnjge24i15pulv8kf2estrrl43en.apps.googleusercontent.com.json"
        flow = InstalledAppFlow.from_client_secrets_file(client_secrets_file, scopes)
        credentials = flow.run_local_server(port=8080)
        with open(credentials_file, "wb") as f:
            pickle.dump(credentials, f)
    return credentials

def get_video_details(youtube, video_id):
    request = youtube.videos().list(
        part="snippet",
        id=video_id
    )
    response = request.execute()

    if "items" in response and len(response["items"]) > 0:
        video = response["items"][0]
        video_description = video["snippet"]["description"]
        video_tags = video["snippet"]["tags"] if "tags" in video["snippet"] else []
        video_published_at = video["snippet"]["publishedAt"]
        return video_description, video_tags, video_published_at

    return "", [], ""

def export_to_excel(video_data):
    wb = Workbook()
    ws = wb.active

    headers = ["Video URL", "Video Title", "Video Description", "Video Tags", "Published At"]
    ws.append(headers)

    for video in video_data:
        video_url = video["url"]
        video_title = video["title"]
        video_description = video["description"]
        video_tags = ", ".join(video["tags"])
        video_published_at = video["published_at"]
        ws.append([video_url, video_title, video_description, video_tags, video_published_at])

    wb.save(output_file)
    print(f"Video data exported to {output_file}")

def main():
    credentials = authenticate()
    youtube = googleapiclient.discovery.build("youtube", "v3", credentials=credentials)

    video_data = []

    next_page_token = None
    while True:
        request = youtube.search().list(
            part="snippet",
            channelId="UCgnhFIfZXi_TUvcAR9cBx_A",
            maxResults=500,
            publishedAfter="2013-01-01T00:00:00Z",  # Filter videos published after January 1, 2022
            publishedBefore="2023-08-31T23:59:59Z",  # Filter videos published before Aug 31, 2023
            pageToken=next_page_token
        )
        response = request.execute()

        for item in response["items"]:
            if "id" in item and "videoId" in item["id"]:
                video_url = f"https://www.youtube.com/watch?v={item['id']['videoId']}"
                video_title = item["snippet"]["title"]
                video_description, video_tags, video_published_at = get_video_details(youtube, item['id']['videoId'])
                video_data.append({
                    "url": video_url,
                    "title": video_title,
                    "description": video_description,
                    "tags": video_tags,
                    "published_at": video_published_at
                })

        next_page_token = response.get("nextPageToken")
        if not next_page_token:
            break

    video_data = sorted(video_data, key=lambda x: x["url"])  # Sort video data by URL in ascending order
    export_to_excel(video_data)

if __name__ == "__main__":
    main()