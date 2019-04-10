"""Python console app with device flow authentication."""
# Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
# See LICENSE in the project root for license information.
import pprint

import config

from helpers import api_endpoint, device_flow_session, profile_photo, \
    send_mail, sharing_link, upload_file

def sendmail_sample(session):
    """Send email from authenticated user.

    session = requests.Session() instance with a valid access token for
              Microsoft Graph in its default HTTP headers

    This sample retrieves the user's profile photo, uploads it to OneDrive,
    creates a view-only sharing link for the photo, and sends an email
    with the photo attached.

    The code in this function includes many print statements to provide
    information about which endpoints are being called and the status and
    size of Microsoft Graph responses. This information is helpful for
    understanding how the sample works with Graph, but would not be included
    in a typical production application.
    """

    print('\nGet user profile ---------> https://graph.microsoft.com/beta/me')
    user_profile = session.get(api_endpoint('me'))
    print(28*' ' + f'<Response [{user_profile.status_code}]>', f'bytes returned: {len(user_profile.text)}\n')
    if not user_profile.ok:
        pprint.pprint(user_profile.json()) # display error
        return
    user_data = user_profile.json()
    email = user_data['mail']
    display_name = user_data['displayName']

    print(f'Your name ----------------> {display_name}')
    print(f'Your email ---------------> {email}')
    email_to = input(f'Send-to (ENTER=self) -----> ') or email

    print('\nGet profile photo --------> https://graph.microsoft.com/beta/me/photo/$value')
    photo, photo_status_code, _, profile_pic = profile_photo(session, save_as='me')
    print(28*' ' + f'<Response [{photo_status_code}]>',
          f'bytes returned: {len(photo)}, saved as: {profile_pic}')
    if not 200 <= photo_status_code <= 299:
        return

    print(f'Upload to OneDrive ------->',
          f'https://graph.microsoft.com/beta/me/drive/root/children/{profile_pic}/content')
    upload_response = upload_file(session, filename=profile_pic)
    print(28*' ' + f'<Response [{upload_response.status_code}]>')
    if not upload_response.ok:
        pprint.pprint(upload_response.json()) # show error message
        return

    print('Create sharing link ------>',
          'https://graph.microsoft.com/beta/me/drive/items/{id}/createLink')
    response, link_url = sharing_link(session, item_id=upload_response.json()['id'])
    print(28*' ' + f'<Response [{response.status_code}]>',
          f'bytes returned: {len(response.text)}')
    if not response.ok:
        pprint.pprint(response.json()) # show error message
        return

    print('Send mail ---------------->',
          'https://graph.microsoft.com/beta/me/microsoft.graph.sendMail')
    with open('email.html') as template_file:
        template = template_file.read().format(name=display_name, link_url=link_url)

    send_response = send_mail(session=session,
                              subject='email from Microsoft Graph console app',
                              recipients=email_to.split(';'),
                              body=template,
                              attachments=[profile_pic])
    print(28*' ' + f'<Response [{send_response.status_code}]>')
    if not send_response.ok:
        pprint.pprint(send_response.json()) # show error message

if __name__ == '__main__':
    GRAPH_SESSION = device_flow_session(config.CLIENT_ID)
    if GRAPH_SESSION:
        sendmail_sample(GRAPH_SESSION)
