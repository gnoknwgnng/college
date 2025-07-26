import streamlit as st
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
import ssl
from groq import Groq # Import Groq library

# --- Groq API Configuration ---
# IMPORTANT: For production, use Streamlit Secrets or environment variables for API keys.
GROQ_API_KEY = "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" # <<< Your Groq API Key
GROQ_MODEL_NAME = "llama-3.1-8b-instant"

client = Groq(api_key=GROQ_API_KEY)

def send_email(sender_email, sender_password, recipient_email, subject, html_body, smtp_server, smtp_port, images=None):
    try:
        msg = MIMEMultipart('related')
        msg['Subject'] = subject
        msg['From'] = sender_email
        msg['To'] = recipient_email

        # Attach HTML part
        msg_html = MIMEText(html_body, 'html')
        msg.attach(msg_html)

        # Attach images
        if images:
            for i, uploaded_image in enumerate(images):
                image_data = uploaded_image.read()
                image_name = uploaded_image.name
                image_type = uploaded_image.type.split('/')[1] # e.g., 'png' from 'image/png'
                # Create a Content-ID based on filename or an index
                cid = f'image_cid_{i}'
                
                img = MIMEImage(image_data, _subtype=image_type, name=image_name)
                img.add_header('Content-ID', f'<{cid}>')
                msg.attach(img)

        context = ssl.create_default_context()

        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls(context=context)
            server.login(sender_email, sender_password)
            server.send_message(msg)
        return True, ""
    except Exception as e:
        return False, str(e)

# Function to get AI response
def get_ai_response(prompt):
    try:
        chat_completion = client.chat.completions.create(
            messages=[
                {
                    "role": "user",
                    "content": prompt,
                }
            ],
            model=GROQ_MODEL_NAME,
        )
        return chat_completion.choices[0].message.content
    except Exception as e:
        return f"Error generating AI response: {e}"

st.set_page_config(layout="wide")
st.title("Excel Email Sender with AI Assistant")

st.write("Upload an Excel file with email addresses, embed images, and use AI to craft personalized emails.")

# Input fields
excel_file = st.file_uploader("Upload your Excel file (.xlsx)", type=["xlsx"])
email_column = st.text_input("Enter the name of the column containing email addresses:", "Email")

st.subheader("Sender Details")
sender_email = st.text_input("Your Email Address:", "")
sender_password = st.text_input("Your Password (or App Password for Gmail):", type="password")
st.info("**For Gmail users:** If you have 2-Step Verification enabled, you'll likely need an App Password. "
        "Go to [Google Account Security](https://myaccount.google.com/security) -> App passwords to generate one.")
# Hardcoded SMTP Server and Port
smtp_server = 'smtp.gmail.com'
smtp_port = 587

st.subheader("AI Email Content Generator")
# AI Chatbot Integration
ai_prompt = st.text_area("Tell the AI what you want in your email (e.g., 'Write a formal email about a meeting. Include a subject and body.'):", height=100)
if st.button("Generate Email Content with AI"):
    if ai_prompt:
        with st.spinner("Generating response..."):
            # Instruct AI to provide both subject and body with clear labels
            full_prompt = "Please provide an email subject and body based on the following: " + ai_prompt + "\n\nFormat your response as: \nSubject: [Your Subject]\n\nBody: [Your Body Content]"
            ai_response_raw = get_ai_response(full_prompt)
            
            # Parse the AI's response
            subject_line = ""
            body_content = ""
            
            if "Subject: " in ai_response_raw and "Body: " in ai_response_raw:
                try:
                    subject_start = ai_response_raw.find("Subject: ") + len("Subject: ")
                    body_start = ai_response_raw.find("Body: ") + len("Body: ")
                    
                    subject_end = ai_response_raw.find("\n\nBody: ", subject_start)
                    if subject_end == -1: # Fallback if "Body: " isn't preceded by two newlines
                        subject_end = ai_response_raw.find("\nBody: ", subject_start)
                    
                    if subject_end != -1:
                        subject_line = ai_response_raw[subject_start:subject_end].strip()
                        body_content = ai_response_raw[body_start:].strip()
                    else:
                        st.warning("Could not parse subject and body from AI response. Please ensure AI formats output correctly.")
                        body_content = ai_response_raw # Fallback to show raw AI response
                except Exception as e:
                    st.error(f"Error parsing AI response: {e}")
                    body_content = ai_response_raw # Fallback to show raw AI response
            else:
                st.warning("AI response did not contain clear 'Subject:' and 'Body:' labels. Using raw response as body.")
                body_content = ai_response_raw # Fallback to show raw AI response

            # Update Streamlit state for both fields
            st.session_state['generated_email_subject'] = subject_line
            st.session_state['generated_email_body'] = body_content
            st.success("AI email content generated!")
            
            # Rerun the app to update the text areas with new values
            st.rerun()

    else:
        st.warning("Please enter a prompt for the AI.")

# Use st.session_state to persist the generated subject
initial_subject = st.session_state.get('generated_email_subject', "")
subject = st.text_input("Email Subject - Editable:", value=initial_subject)

# Use st.session_state to persist the generated body
initial_body_text = st.session_state.get('generated_email_body', """
Dear recipient,

This is a test email with an embedded image.

Best regards,
Your Name
""")

body_plain_text = st.text_area(
    "Email Body (Plain Text) - Editable:", 
    value=initial_body_text, 
    height=300
)

# Image Uploader
st.subheader("Embed Images")
uploaded_images = st.file_uploader("Upload images to embed (PNG, JPG, JPEG, GIF)", type=["png", "jpg", "jpeg", "gif"], accept_multiple_files=True)

# Display uploaded images (for reference, though user won't directly type CID now)
if uploaded_images:
    st.subheader("Uploaded Images (will be embedded automatically):")
    for i, uploaded_file in enumerate(uploaded_images):
        st.write(f"- **{uploaded_file.name}**") 
        st.image(uploaded_file, width=100)

if st.button("Send Emails"):
    if excel_file is not None:
        if not sender_email or not sender_password or not subject or not body_plain_text:
            st.warning("Please fill in all sender details, subject, and email body.")
        else:
            try:
                df = pd.read_excel(excel_file)

                if email_column not in df.columns:
                    st.error(f"Error: Column '{email_column}' not found in the Excel file.")
                else:
                    email_addresses = df[email_column].dropna().tolist()

                    if not email_addresses:
                        st.warning("No email addresses found in the specified column.")
                    else:
                        st.info(f"Found {len(email_addresses)} email addresses. Starting to send emails...")
                        
                        progress_bar = st.progress(0)
                        status_text = st.empty()
                        total_emails = len(email_addresses)

                        # Generate HTML body with embedded images
                        generated_html_body = body_plain_text.replace("\n", "<br>") # Convert newlines to HTML breaks
                        if uploaded_images:
                            for i, uploaded_file in enumerate(uploaded_images):
                                cid = f'image_cid_{i}'
                                generated_html_body += f'<br><img src="cid:{cid}" alt="Embedded Image {i}"><br>' # Append images to body

                        for i, recipient_email in enumerate(email_addresses):
                            success, error_msg = send_email(
                                sender_email, 
                                sender_password, 
                                recipient_email, 
                                subject, 
                                generated_html_body, 
                                smtp_server, 
                                smtp_port,
                                images=uploaded_images
                            )
                            if success:
                                status_text.write(f"Email sent successfully to {recipient_email}")
                            else:
                                st.error(f"Failed to send email to {recipient_email}: {error_msg}")
                            progress_bar.progress((i + 1) / total_emails)
                        st.success("Email sending process completed.")

            except Exception as e:
                st.error(f"An error occurred while processing the Excel file: {e}")
    else:
        st.warning("Please upload an Excel file.") 