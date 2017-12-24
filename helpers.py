"""helper functions for Microsoft Graph"""
# Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
# See LICENSE in the project root for license information.
import base64
import mimetypes
import os
import urllib
import webbrowser

from adal import AuthenticationContext
import pyperclip
import requests

import config

def api_endpoint(url):
    """Convert a relative path such as /me/photo/$value to a full URI based
    on the current RESOURCE and API_VERSION settings in config.py.
    """
    if urllib.parse.urlparse(url).scheme in ['http', 'https']:
        return url # url is already complete
    return urllib.parse.urljoin(f'{config.RESOURCE}/{config.API_VERSION}/',
                                url.lstrip('/'))

def device_flow_session(client_id, auto=False):
    """Obtain an access token from Azure AD (via device flow) and create
    a Requests session instance ready to make authenticated calls to
    Microsoft Graph.

    client_id = Application ID for registered "Azure AD only" V1-endpoint app
    auto      = whether to copy device code to clipboard and auto-launch browser

    Returns Requests session object if user signed in successfully. The session
    includes the access token in an Authorization header.

    User identity must be an organizational account (ADAL does not support MSAs).
    """
    ctx = AuthenticationContext(config.AUTHORITY_URL, api_version=None)
    device_code = ctx.acquire_user_code(config.RESOURCE,
                                        client_id)

    # display user instructions
    if auto:
        pyperclip.copy(device_code['user_code']) # copy user code to clipboard
        webbrowser.open(device_code['verification_url']) # open browser
        print(f'The code {device_code["user_code"]} has been copied to your clipboard, '
              f'and your web browser is opening {device_code["verification_url"]}. '
              'Paste the code to sign in.')
    else:
        print(device_code['message'])

    token_response = ctx.acquire_token_with_device_code(config.RESOURCE,
                                                        device_code,
                                                        client_id)
    if not token_response.get('accessToken', None):
        return None

    session = requests.Session()
    session.headers.update({'Authorization': f'Bearer {token_response["accessToken"]}',
                            'SdkVersion': 'sample-python-adal',
                            'x-client-SKU': 'sample-python-adal'})
    return session

def profile_photo(session, *, user_id='me', save_as=None):
    """Get profile photo, and optionally save a local copy.

    session = requests.Session() instance with Graph access token
    user_id = Graph id value for the user, or 'me' (default) for current user
    save_as = optional filename to save the photo locally. Should not include an
              extension - the extension is determined by photo's content type.

    Returns a tuple of the photo (raw data), HTTP status code, content type, saved filename.
    """

    endpoint = 'me/photo/$value' if user_id == 'me' else f'users/{user_id}/$value'
    photo_response = session.get(api_endpoint(endpoint),
                                 stream=True)
    photo_status_code = photo_response.status_code
    if photo_response.ok:
        photo = photo_response.raw.read()
        # note we remove /$value from endpoint to get metadata endpoint
        metadata_response = session.get(api_endpoint(endpoint[:-7]))
        content_type = metadata_response.json().get('@odata.mediaContentType', '')
    else:
        photo = ''
        content_type = ''

    if photo and save_as:
        extension = content_type.split('/')[1]
        filename = save_as + '.' + extension
        with open(filename, 'wb') as fhandle:
            fhandle.write(photo)
    else:
        filename = ''

    return (photo, photo_status_code, content_type, filename)

def send_mail(session, *, subject, recipients, body='', content_type='HTML',
              attachments=None):
    """Send email from current user.

    session      = requests.Session() instance with Graph access token
    subject      = email subject (required)
    recipients   = list of recipient email addresses (required)
    body         = body of the message
    content_type = content type (default is 'HTML')
    attachments  = list of file attachments (local filenames)

    Returns the response from the POST to the sendmail API.
    """

    # Create recipient list in required format.
    recipient_list = [{'EmailAddress': {'Address': address}}
                      for address in recipients]

    # Create list of attachments in required format.
    attached_files = []
    if attachments:
        for filename in attachments:
            b64_content = base64.b64encode(open(filename, 'rb').read())
            mime_type = mimetypes.guess_type(filename)[0]
            mime_type = mime_type if mime_type else ''
            attached_files.append( \
                {'@odata.type': '#microsoft.graph.fileAttachment',
                 'ContentBytes': b64_content.decode('utf-8'),
                 'ContentType': mime_type,
                 'Name': filename})

    # Create email message in required format.
    email_msg = {'Message': {'Subject': subject,
                             'Body': {'ContentType': content_type, 'Content': body},
                             'ToRecipients': recipient_list,
                             'Attachments': attached_files},
                 'SaveToSentItems': 'true'}

    # Do a POST to Graph's sendMail API and return the response.
    return session.post(api_endpoint('me/microsoft.graph.sendMail'),
                        headers={'Content-Type': 'application/json'},
                        json=email_msg)

def sharing_link(session, *, item_id, link_type='view'):
    """Get a sharing link for an item in OneDrive.

    session   = requests.Session() instance with Graph access token
    item_id   = the id of the DriveItem (the target of the link)
    link_type = 'view' (default), 'edit', or 'embed' (OneDrive Personal only)

    Returns a tuple of the response object and the sharing link.
    """
    endpoint = f'me/drive/items/{item_id}/createLink'
    response = session.post(api_endpoint(endpoint),
                            headers={'Content-Type': 'application/json'},
                            json={'type': link_type})

    if response.ok:
        # status 201 = link created, status 200 = existing link returned
        return (response, response.json()['link']['webUrl'])
    return (response, '')

def upload_file(session, *, filename, folder=None):
    """Upload a file to OneDrive for Business.

    session  = requests.Session() instance with Graph access token
    filename = local filename; may include a path
    folder   = destination subfolder/path in OneDrive for Business
               None (default) = root folder

    File is uploaded and the response object is returned.
    If file already exists, it is overwritten.
    If folder does not exist, it is created.

    API documentation:
    https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/api/driveitem_put_content
    """
    fname_only = os.path.basename(filename)

    # create the Graph endpoint to be used
    if folder:
        # create endpoint for upload to a subfolder
        endpoint = f'me/drive/root:/{folder}/{fname_only}:/content'
    else:
        # create endpoint for upload to drive root folder
        endpoint = f'me/drive/root/children/{fname_only}/content'

    content_type, _ = mimetypes.guess_type(fname_only)
    with open(filename, 'rb') as fhandle:
        file_content = fhandle.read()

    return session.put(api_endpoint(endpoint),
                       headers={'content-type': content_type},
                       data=file_content)
