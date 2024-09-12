# Mass Email Sender & Certificate Generator

This Python project allows users to send mass emails with personalized certificates attached. It is especially useful for events, courses, or any occasion where bulk certificate generation and distribution are required.

## Features

- **Mass Email Sending**: Send emails to multiple recipients at once.
- **Personalized Certificates**: Automatically generate certificates with personalized details (e.g., name, event) using a predefined certificate template.
- **Attachment Support**: Attach the generated certificates to the emails.
- **SMTP Integration**: Send emails via an SMTP server (Gmail, Outlook, etc.).
- **CSV Support**: Load recipient details from a CSV file.
  
## Prerequisites

Make sure you have the following installed:

- **Python 3.x**
- **Pillow**: To handle image manipulation for certificate generation.
- **smtplib**: For sending emails (part of Python standard library).
- **email.mime**: For constructing the email messages (part of Python standard library).
- **tkinter**: For creating the graphical user interface (optional, if GUI is used).

You can install any missing libraries using:

```bash
pip install Pillow
```

## How It Works

1. **Load Recipients**: Upload a CSV file containing recipient details (name, email, etc.).
2. **Generate Certificates**: Personalize and generate certificates based on the uploaded data.
3. **Send Emails**: Send the certificates via email to all recipients, using the provided SMTP configuration.

## Getting Started

### 1. Clone the Repository

```bash
git clone https://github.com/yourusername/mass-email-sender.git
cd mass-email-sender
```

### 2. Install Dependencies

If you're using a `requirements.txt` file, install the dependencies with:

```bash
pip install -r requirements.txt
```

### 3. CSV File Structure

The CSV file should contain the following columns:

| Name       | Email            | Other Details (optional) |
|------------|------------------|--------------------------|
| John Doe   | john@example.com  | Event Date               |
| Jane Smith | jane@example.com  |                          |

### 4. Configuring the Certificate Template

- Place your certificate template in the project folder.
- The template should have placeholders where the recipient's name and details will be added.
  
Modify the certificate generation function to match the placeholder positions and fonts in the template.

```python
from PIL import Image, ImageDraw, ImageFont

def generate_certificate(recipient_name):
    template = Image.open("certificate_template.png")
    draw = ImageDraw.Draw(template)
    
    # Load a font
    font = ImageFont.truetype("arial.ttf", 60)
    
    # Add recipient name to the certificate
    draw.text((x_position, y_position), recipient_name, font=font, fill="black")
    
    # Save the certificate with the recipient's name
    template.save(f"certificates/{recipient_name}_certificate.png")
```

### 5. SMTP Configuration

Update the SMTP settings with your email provider's credentials.

Example for Gmail:

```python
smtp_server = "smtp.gmail.com"
smtp_port = 587
smtp_user = "your-email@gmail.com"
smtp_password = "your-app-password"
```

If using Gmail, you may need to enable "less secure apps" or use an app-specific password.

### 6. Running the Program

You can either run the script with a command-line interface (CLI) or a graphical user interface (GUI) built with Tkinter.

#### CLI Usage

```bash
python mass_email_sender.py --csv recipients.csv --subject "Your Certificate" --body "Please find your certificate attached." --template certificate_template.png
```

#### GUI Usage (if implemented)

Run the GUI by executing:

```bash
python email_sender_gui.py
```

## How to Use

1. **Load CSV**: Load the CSV file with recipient details.
2. **Generate Certificates**: For each recipient, a personalized certificate will be generated.
3. **Send Emails**: The program will send an email to each recipient with their certificate attached.

## Future Improvements

- **Multiple Email Providers**: Add support for more email providers and better error handling for SMTP issues.
- **Customizable Templates**: Allow users to upload and customize certificate templates from the GUI.
- **Bulk Scheduling**: Schedule emails to be sent at a later date.

## Contributing

If you would like to contribute to this project, feel free to submit a pull request. Bug reports and suggestions are welcome.

## License

This project is licensed under the MIT License.

---

This version of the `README.md` focuses on a mass email sender and certificate generator. It provides clear instructions for setting up, configuring SMTP, and generating personalized certificates. Let me know if you'd like further adjustments!
