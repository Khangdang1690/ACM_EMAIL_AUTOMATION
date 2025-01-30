import win32com.client
import pandas as pd

def send_outlook_email(subject, body, recipient_name, recipient_email, image_path=None, attachment_paths=None):
    # Create an Outlook application instance
    outlook = win32com.client.Dispatch('Outlook.Application')
    mail = outlook.CreateItem(0)  # 0: olMailItem (Standard email item)

    # Customize the body to include the professor's name
    body = body.replace("{name}", recipient_name)

    # Set up the subject and body of the email
    mail.Subject = subject
    mail.HTMLBody = body

    # Add recipient(s)
    mail.To = recipient_email

    # Add the image if provided (embed it inline)
    if image_path:
        embed_image(mail, image_path)

    # Attach files if provided
    if attachment_paths:
        for attachment_path in attachment_paths:
            attach_file(mail, attachment_path)

    # Send the email
    mail.Send()

def embed_image(mail, image_path):
    # Open the image file
    with open(image_path, 'rb') as img_file:
        img_data = img_file.read()

    # Attach the image as inline content
    attachment = mail.Attachments.Add(image_path)
    attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", 'logo_image')
    
    # This 'Content-ID' should match the reference in the HTML <img> tag
    # Note: "logo_image" here matches the 'src="cid:logo_image"' in the HTML email

def attach_file(mail, attachment_path):
    # Attach the file
    attachment = mail.Attachments.Add(attachment_path)
    # Optional: You can set additional properties for the attachment if needed
    # attachment.DisplayName = "Custom Display Name"

def main():
    # Email content
    subject = "Request for Support: ACM AI Convention 2025"
    body = """
    <html>
    <body>
    <p>Dear Dr. {name},</p>

    <p>I hope this email finds you well. My name is Khang Dang, and I serve as Marketing Shadow for the Association for Computing Machinery (ACM) Chapter at USF.</p>

    <p>We are thrilled to host the <strong>USF ACM AI Convention 2025 on February 8th</strong>, a full-day event featuring keynote speakers, hands-on workshops, project showcases, and networking opportunities. Our goal is to encourage students across disciplines to explore how AI is shaping industries such as education, healthcare, engineering, and business.</p>

    <p>We would greatly appreciate your assistance in promoting this event by sending an email announcement to all students in your department. This would be instrumental in helping us reach a wider audience and encouraging participation in this exciting and interdisciplinary event.</p>

    <p>For your convenience, we have attached the event flyer and itinerary, along with a short description that can be included in the email:</p>
    <blockquote>
    <p><strong>USF ACM: AI Convention 2025</strong><br>
    <strong>Date: February 8, 2025 - ENB</strong></p>

    <p>🚀 Have you ever wondered how AI can shape the career path that you’re pursuing? You may have an idea about it, but have you actually seen it with your own eyes? For example, how can a normal person leverage AI to create a piece of art that is better than what an experienced person can generate? Or how doctors and biotechnologists are using AI to discover big breakthroughs and save thousands of lives. Or, how you can find your dream job with AI. All that can be seen on February 8, 2025 at ENB, and you just need to show up and watch the biggest convention displayed in front of your eyes.</p>
    </blockquote>

    <p>Your support would make a significant difference in engaging students with the opportunities and insights offered at the AI Convention. Please feel free to reach out if you have any questions or need additional materials.</p>

    <p>Thank you for considering this request, and I look forward to your support in making this event a success!</p>

    <p>Sincerely,<br>
    <strong>Khang Dang</strong><br>
    B.S. Computer Science | University of South Florida<br>
    Marketing Shadow | Association for Computing Machinery, USF <br>
    Resources: Instagram | Linktree | Website</p>

    <!-- Insert Image Below the Text with Adjusted Size -->
    <img src="cid:logo_image" alt="ACM Logo" width="300" height="auto" />
    </body>
    </html>
    """

    # Load Excel file with professor names and emails
    df = pd.read_excel("professors.xlsx")  # Replace with the path to your Excel file

    # List of file paths to attach (add the paths of the two files you want to attach)
    attachment_paths = [
        r"C:\Users\khang\OneDrive\Desktop\email_automation\event_flyer.webp",  # Example file path
        r"C:\Users\khang\OneDrive\Desktop\email_automation\event_itinerary.pdf"  # Example file path
    ]

    # Loop through each row in the Excel file
    for _, row in df.iterrows():
        recipient_name = row['Name']
        recipient_email = row['Email']

        # Path to the logo image
        image_path = r"C:\Users\khang\OneDrive\Desktop\email_automation\logo.png"  # Replace with your logo file path

        # Send email
        send_outlook_email(subject, body, recipient_name, recipient_email, image_path, attachment_paths)

        print(f"Email sent to {recipient_name} ({recipient_email})")

if __name__ == "__main__":
    main()
