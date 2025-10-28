import pywhatkit

# Replace with your number and message
pywhatkit.sendwhatmsg_instantly(
    phone_no="+9191",
    message="Hello! This is a free WhatsApp message from Python.",
    wait_time=10,  # seconds to wait before sending the message
    tab_close=True
)
