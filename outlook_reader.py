import win32com.client
# import datetime


def connect_to_outlook():
    """Connect to Outlook application"""
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        return namespace
    except Exception as e:
        print(f"Error connecting to Outlook: {e}")
        return None


def get_inbox(namespace):
    """Get the inbox folder"""
    try:
        inbox = namespace.GetDefaultFolder(6)
        return inbox
    except Exception as e:
        print(f"Error accessing inbox: {e}")
        return None


def read_emails(inbox, num_emails=10):
    """Read the specified number of emails from inbox"""
    try:
        messages = inbox.Items
        messages.Sort("[ReceivedTime]", True)

        for i, message in enumerate(messages):
            if i >= num_emails:
                break

            print("\n" + "="*50)
            print(f"Subject: {message.Subject}")
            print(f"Sender: {message.SenderName}")
            print(f"Received: {message.ReceivedTime}")
            print(f"Body Preview: {message.Body[:100]}...")

    except Exception as e:
        print(f"Error reading emails: {e}")


def main():
    # Connect to Outlook
    namespace = connect_to_outlook()
    if not namespace:
        return

    # Get inbox
    inbox = get_inbox(namespace)
    if not inbox:
        return

    # Read last 5 emails
    print("Reading the last 5 emails from your inbox...")
    read_emails(inbox, 5)


if __name__ == "__main__":
    main()
