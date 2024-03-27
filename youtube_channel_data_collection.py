import os
import pickle
from google_auth_oauthlib.flow import InstalledAppFlow
import googleapiclient.discovery
import googleapiclient.errors
from openpyxl import Workbook

# Set the proxy server details
proxy_server = "http://Your_address:Your_port_number"

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
        client_secrets_file = "client_secret.json"
        flow = InstalledAppFlow.from_client_secrets_file(client_secrets_file, scopes)
        credentials = flow.run_local_server(port=8080)
        with open(credentials_file, "wb") as f:
            pickle.dump(credentials, f)
    return credentials

def get_video_details(youtube, video_id):
    request = youtube.videos().list(
        part="snippet,statistics",
        id=video_id
    )
    response = request.execute()

    if "items" in response and len(response["items"]) > 0:
        video = response["items"][0]
        video_description = video["snippet"]["description"]
        video_tags = video["snippet"]["tags"] if "tags" in video["snippet"] else []
        video_published_at = video["snippet"]["publishedAt"]
        video_views = video["statistics"]["viewCount"] if "viewCount" in video["statistics"] else 0
        return video_description, video_tags, video_published_at, video_views

    return "", [], "", 0

def export_to_excel(video_data):
    wb = Workbook()
    ws = wb.active

    headers = ["Video URL", "Video Title", "Video Description", "Video Tags", "Published At", "Views"]
    ws.append(headers)

    for video in video_data:
        video_url = video["url"]
        video_title = video["title"]
        video_description = video["description"]
        video_tags = ", ".join(video["tags"])
        video_published_at = video["published_at"]
        video_views = video["views"]
        ws.append([video_url, video_title, video_description, video_tags, video_published_at, video_views])

    wb.save(output_file)
    print(f"Video data exported to {output_file}")

def main():
    credentials = authenticate()
    youtube = googleapiclient.discovery.build("youtube", "v3", credentials=credentials)

    video_data = []
    processed_video_ids = set()  # Set to store processed video IDs

    next_page_token = None
    while True:
        request = youtube.search().list(
            part="snippet",
            channelId="xxxxxxxxxxxx",  # Here is the channel id
            maxResults=500,
            publishedAfter="2020-01-01T00:00:00Z",  # Filter videos published after January 1, 2022
            publishedBefore="2024-03-21T23:59:59Z",  # Filter videos published before Aug 31, 2023
            pageToken=next_page_token
        )
        response = request.execute()

        for item in response["items"]:
            if "id" in item and "videoId" in item["id"]:
                video_id = item['id']['videoId']
                if video_id in processed_video_ids:
                    continue  # Skip processing if video ID is already processed

                video_url = f"https://www.youtube.com/watch?v={video_id}"
                video_title = item["snippet"]["title"]
                video_description, video_tags, video_published_at, video_views = get_video_details(youtube, video_id)
                video_data.append({
                    "url": video_url,
                    "title": video_title,
                    "description": video_description,
                    "tags": video_tags,
                    "published_at": video_published_at,
                    "views": video_views
                })

                processed_video_ids.add(video_id)  # Add video ID to processed set

        next_page_token = response.get("nextPageToken")
        if not next_page_token:
            break

    video_data = sorted(video_data, key=lambda x: x["url"])  # Sort video data by URL in ascending order
    export_to_excel(video_data)

if __name__ == "__main__":
    main()
