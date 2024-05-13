import requests
def send_message_in_discord_channel(webhook_url, message_content, logger):
    # send the respective employee details to the manager to approve the leave request
    data = {
        "content": message_content
    }
    response = requests.post(webhook_url, data=data)
    # Check whether the response send or not
    if response.status_code == 204:
        logger.info(f"{response.text} \n Message sent successfully!")
    else:
        logger.info(f"Failed to send message. Status code: {response.status_code} {response.text}")
