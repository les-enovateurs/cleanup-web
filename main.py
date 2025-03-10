import os
import email
import re
import csv
import zipfile
import xml.etree.ElementTree as ET
import sys
import traceback
from email.utils import parseaddr, parsedate_to_datetime
from datetime import datetime
from collections import defaultdict, Counter
import argparse
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                            QLabel, QLineEdit, QPushButton, QProgressBar, QTextEdit,
                            QFileDialog, QMessageBox, QFrame, QRadioButton, QButtonGroup,
                            QGroupBox, QGridLayout)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QUrl
from PyQt6.QtGui import QPixmap, QDesktopServices

class EmailExportParser:
    def __init__(self, export_dir=None):
        """
        Initialize the parser with the directory containing the email export files.
        
        Args:
            export_dir (str, optional): Path to the directory containing email export files
        """
        self.export_dir = export_dir
        self.commercial_emails = defaultdict(list)
        self.unsubscribe_urls = {}
        self.status_callback = None
        self.progress_callback = None
        self.total_emails_processed = 0
        self.commercial_emails_found = 0
        self.recipient_emails = Counter()  # Track recipient email addresses

        # Commercial keywords in English and French
        self.commercial_keywords = [
            # English keywords
            'unsubscribe', 'marketing', 'newsletter', 'subscription', 'promotional',
            'offer', 'discount', 'sale', 'deal', 'advertisement', 'sponsored',
            'coupon', 'promo', 'limited time', 'special offer',

            # French keywords
            'désabonner', 'désabonnement', 'marketing', 'infolettre', 'bulletin',
            'abonnement', 'promotionnel', 'promotion', 'offre', 'remise', 'solde',
            'réduction', 'vente', 'publicité', 'parrainé', 'coupon', 'promo',
            'durée limitée', 'offre spéciale', 'newsletter', 'annulation'
        ]

    def set_status_callback(self, callback):
        """Set a callback function to receive status updates."""
        self.status_callback = callback

    def set_progress_callback(self, callback):
        """Set a callback function to receive progress updates."""
        self.progress_callback = callback

    def update_status(self, message):
        """Update status with a message."""
        if self.status_callback:
            self.status_callback(message)
        else:
            print(message)

    def update_progress(self, current, total=None):
        """Update progress."""
        if self.progress_callback:
            self.progress_callback(current, total)

    def parse_export(self):
        """Parse the email export files and identify commercial emails."""
        if not self.export_dir:
            raise ValueError("Export directory not set")

        # Check if the directory exists
        if not os.path.exists(self.export_dir):
            raise FileNotFoundError(f"Export directory {self.export_dir} not found")

        self.update_status(f"Starting to process files in: {self.export_dir}")
        self.commercial_emails.clear()
        self.unsubscribe_urls.clear()
        self.total_emails_processed = 0
        self.commercial_emails_found = 0

        # Get list of all files to process for progress estimation
        all_files = []
        for root, _, files in os.walk(self.export_dir):
            for file in files:
                ext = os.path.splitext(file)[1].lower()
                if ext in ['.mbox', '.eml', '.pst', '.ost', '.msg', '.xml', '.zip']:
                    all_files.append(os.path.join(root, file))

        total_files = len(all_files)
        self.update_status(f"Found {total_files} email files to process")

        # Process various email export formats
        files_processed = 0
        for file_path in all_files:
            try:
                file_name = os.path.basename(file_path)
                file_ext = os.path.splitext(file_name)[1].lower()

                # Update progress
                files_processed += 1
                self.update_progress(files_processed, total_files)
                self.update_status(f"Processing file {files_processed}/{total_files}: {file_name}")

                # Detect file type and process accordingly
                if file_ext == '.mbox':
                    self._process_mbox(file_path)
                elif file_ext == '.eml':
                    self._process_eml(file_path)
                elif file_ext in ['.pst', '.ost']:
                    self._process_outlook_database(file_path)
                elif file_ext == '.msg':
                    self._process_outlook_msg(file_path)
                elif file_ext == '.xml' and 'outlook' in file_name.lower():
                    self._process_outlook_xml(file_path)
                elif file_ext == '.zip' and 'outlook' in file_name.lower():
                    self._process_outlook_zip(file_path)
            except Exception as e:
                self.update_status(f"Error processing {file_path}: {str(e)}")

        self.update_status(f"Completed processing {files_processed} files")
        self.update_status(f"Processed {self.total_emails_processed} emails total")
        self.update_status(f"Found {self.commercial_emails_found} commercial emails from {len(self.commercial_emails)} domains")

        return self.commercial_emails

    def _process_mbox(self, mbox_path):
        """Process an .mbox file containing multiple emails."""
        self.update_status(f"Processing Gmail mbox file: {os.path.basename(mbox_path)}")

        # Try to use the mailbox module if available
        try:
            import mailbox
            mbox = mailbox.mbox(mbox_path)
            total_msgs = len(mbox)
            self.update_status(f"Found {total_msgs} messages in mbox file")

            for i, msg in enumerate(mbox):
                if i % 100 == 0:  # Update progress every 100 emails
                    self.update_status(f"Processing email {i+1}/{total_msgs} in {os.path.basename(mbox_path)}")

                self._analyze_email(msg)
                self.total_emails_processed += 1

        except ImportError:
            # Fall back to manual parsing if mailbox module is not available
            self.update_status("Using manual mbox parsing (install 'mailbox' for better performance)")
            try:
                with open(mbox_path, 'rb') as f:
                    content = f.read().decode('utf-8', errors='replace')

                # Split the mbox file into individual messages
                messages = content.split('From ')
                total_msgs = len(messages) - 1  # Skip first empty entry
                self.update_status(f"Found approximately {total_msgs} messages in mbox file")

                for i, message in enumerate(messages[1:]):  # Skip the first empty entry
                    if i % 100 == 0:  # Update progress every 100 emails
                        self.update_status(f"Processing email {i+1}/{total_msgs} in {os.path.basename(mbox_path)}")

                    try:
                        msg = email.message_from_string('From ' + message)
                        self._analyze_email(msg)
                        self.total_emails_processed += 1
                    except Exception as e:
                        self.update_status(f"Error processing message: {e}")
            except Exception as e:
                self.update_status(f"Error processing mbox file: {e}")

    def _process_eml(self, eml_path):
        """Process a single .eml file."""
        self.update_status(f"Processing .eml file: {os.path.basename(eml_path)}")
        try:
            with open(eml_path, 'rb') as f:
                msg = email.message_from_binary_file(f)
                self._analyze_email(msg)
                self.total_emails_processed += 1
        except Exception as e:
            self.update_status(f"Error processing {eml_path}: {e}")

    def _process_outlook_database(self, file_path):
        """
        Process Outlook .pst or .ost files.
        Note: This requires libpff library (pypff) which may not be available in all environments.
        """
        self.update_status(f"Processing Outlook database file: {os.path.basename(file_path)}")
        try:
            # Try to import pypff for processing .pst/.ost files
            import pypff

            pst_file = pypff.file()
            pst_file.open(file_path)
            root_folder = pst_file.get_root_folder()

            # Process all folders recursively
            self._process_outlook_folder(root_folder)

            pst_file.close()

        except ImportError:
            self.update_status("pypff library not available. Cannot process .pst/.ost files directly.")
            self.update_status("Please install with: pip install pypff")
            self.update_status("Or export your Outlook data to .eml or .msg files.")

    def _process_outlook_folder(self, folder):
        """Process an Outlook folder recursively (using pypff)."""
        folder_name = folder.get_name() or "Unknown"
        num_messages = folder.get_number_of_sub_messages()

        if num_messages > 0:
            self.update_status(f"Processing Outlook folder '{folder_name}' with {num_messages} messages")

        # Process all messages in this folder
        for message_index in range(num_messages):
            message = folder.get_sub_message(message_index)

            # Convert pypff message to email.message format
            msg = email.message.EmailMessage()
            msg['Subject'] = message.get_subject() or ''
            msg['From'] = message.get_sender() or ''
            msg['To'] = message.get_recipients() or ''
            msg['Date'] = message.get_delivery_time() or ''

            # Get body
            body = message.get_plain_text_body() or ''
            if body:
                msg.set_content(body)

            html_body = message.get_html_body() or ''
            if html_body:
                msg.add_alternative(html_body, subtype='html')

            self._analyze_email(msg)
            self.total_emails_processed += 1

            if message_index > 0 and message_index % 100 == 0:
                self.update_status(f"Processed {message_index}/{num_messages} messages in folder '{folder_name}'")

        # Process subfolders recursively
        for subfolder_index in range(folder.get_number_of_sub_folders()):
            subfolder = folder.get_sub_folder(subfolder_index)
            self._process_outlook_folder(subfolder)

    def _process_outlook_msg(self, msg_path):
        """Process a single Outlook .msg file."""
        self.update_status(f"Processing Outlook .msg file: {os.path.basename(msg_path)}")
        try:
            # Try to import extract_msg for processing .msg files
            import extract_msg

            outlook_msg = extract_msg.Message(msg_path)

            # Convert to email.message format
            msg = email.message.EmailMessage()
            msg['Subject'] = outlook_msg.subject or ''
            msg['From'] = outlook_msg.sender or ''
            msg['To'] = outlook_msg.to or ''
            msg['Date'] = outlook_msg.date or ''

            # Get body
            body = outlook_msg.body or ''
            if body:
                msg.set_content(body)

            html_body = outlook_msg.htmlBody or ''
            if html_body:
                msg.add_alternative(html_body, subtype='html')

            self._analyze_email(msg)
            self.total_emails_processed += 1

        except ImportError:
            self.update_status("extract_msg library not available. Cannot process .msg files directly.")
            self.update_status("Please install with: pip install extract_msg")

    def _process_outlook_xml(self, xml_path):
        """Process Outlook XML export file."""
        self.update_status(f"Processing Outlook XML file: {os.path.basename(xml_path)}")
        try:
            tree = ET.parse(xml_path)
            root = tree.getroot()

            message_elems = root.findall('.//Message')
            total_msgs = len(message_elems)
            self.update_status(f"Found {total_msgs} messages in XML file")

            for i, message_elem in enumerate(message_elems):
                if i % 100 == 0:  # Update progress every 100 emails
                    self.update_status(f"Processing email {i+1}/{total_msgs} in {os.path.basename(xml_path)}")

                # Create an email message from XML data
                msg = email.message.EmailMessage()

                # Extract fields from XML
                for field in ['Subject', 'From', 'To', 'Date']:
                    elem = message_elem.find(f'.//{field}')
                    if elem is not None and elem.text:
                        msg[field] = elem.text

                # Get body
                body_elem = message_elem.find('.//Body')
                if body_elem is not None and body_elem.text:
                    msg.set_content(body_elem.text)

                self._analyze_email(msg)
                self.total_emails_processed += 1

        except Exception as e:
            self.update_status(f"Error processing Outlook XML file {xml_path}: {e}")

    def _process_outlook_zip(self, zip_path):
        """Process Outlook ZIP export file containing emails."""
        self.update_status(f"Processing Outlook ZIP export: {os.path.basename(zip_path)}")
        try:
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                # List all files in the zip
                file_list = zip_ref.namelist()
                self.update_status(f"Found {len(file_list)} files in ZIP archive")

                # Create a temporary directory for extraction
                temp_dir = os.path.join(os.path.dirname(zip_path), 'temp_outlook_extract')
                os.makedirs(temp_dir, exist_ok=True)

                # Extract files
                self.update_status(f"Extracting files to temporary directory...")
                zip_ref.extractall(temp_dir)

                # Process the extracted files
                extracted_files = []
                for root, _, files in os.walk(temp_dir):
                    for file in files:
                        file_path = os.path.join(root, file)
                        if file.endswith(('.eml', '.msg', '.xml')):
                            extracted_files.append(file_path)

                self.update_status(f"Found {len(extracted_files)} email files in the ZIP archive")

                # Process each file
                for i, file_path in enumerate(extracted_files):
                    file_name = os.path.basename(file_path)
                    self.update_status(f"Processing extracted file {i+1}/{len(extracted_files)}: {file_name}")

                    if file_name.endswith('.eml'):
                        self._process_eml(file_path)
                    elif file_name.endswith('.msg'):
                        self._process_outlook_msg(file_path)
                    elif file_name.endswith('.xml'):
                        self._process_outlook_xml(file_path)

                # Clean up
                self.update_status("Cleaning up temporary files...")
                import shutil
                shutil.rmtree(temp_dir)

        except Exception as e:
            self.update_status(f"Error processing Outlook ZIP file {zip_path}: {e}")

    def _analyze_email(self, msg):
        """Analyze an email message to determine if it's commercial."""
        # Extract basic information
        from_name, from_email = parseaddr(msg.get('From', ''))
        subject = msg.get('Subject', '')
        date_str = msg.get('Date', '')

        # Try to parse the date
        try:
            date = parsedate_to_datetime(date_str)
        except:
            try:
                # Alternative date formats
                date_formats = [
                    "%a, %d %b %Y %H:%M:%S %z",
                    "%d %b %Y %H:%M:%S %z",
                    "%a, %d %b %Y %H:%M:%S",
                    "%d/%m/%Y %H:%M:%S",
                    "%Y-%m-%d %H:%M:%S"
                ]

                for fmt in date_formats:
                    try:
                        date = parsedate_to_datetime(date_str[:31], fmt)
                        break
                    except:
                        continue
                else:
                    date = None
            except:
                date = None

        # Look for unsubscribe links and commercial indicators
        is_commercial = False
        unsubscribe_url = None

        # Check headers for List-Unsubscribe
        if msg.get('List-Unsubscribe'):
            is_commercial = True
            unsubscribe_header = msg.get('List-Unsubscribe')
            url_match = re.search(r'<(https?://[^>]+)>', unsubscribe_header)
            if url_match:
                unsubscribe_url = url_match.group(1)

        # Check email body for commercial indicators and unsubscribe links
        for part in msg.walk():
            content_type = part.get_content_type() if hasattr(part, 'get_content_type') else None

            if content_type in ('text/plain', 'text/html'):
                try:
                    # Get payload (handle different email message types)
                    if hasattr(part, 'get_payload'):
                        if callable(part.get_payload):
                            payload = part.get_payload(decode=True)
                        else:
                            payload = part.get_payload()
                            if isinstance(payload, list):
                                payload = payload[0].get_payload(decode=True)
                            else:
                                payload = part.get_payload(decode=True)
                    else:
                        continue

                    # Decode content
                    if isinstance(payload, bytes):
                        content = payload.decode('utf-8', errors='replace')
                    else:
                        content = str(payload)

                    # Look for commercial keywords (both English and French)
                    for keyword in self.commercial_keywords:
                        if keyword.lower() in content.lower():
                            is_commercial = True
                            break

                    # Look for unsubscribe URLs
                    if not unsubscribe_url:
                        # English unsubscribe patterns
                        patterns = [
                            r'https?://[^\s<>"\']+[^\s<>"\'.,)](?=[^A-Za-z0-9]|$)',
                            r'(?:https?://|www\.)[^\s<>"\']+\bunsubscribe\b[^\s<>"\']*',
                            r'(?:https?://|www\.)[^\s<>"\']+\bopt[ -]?out\b[^\s<>"\']*',

                            # French patterns
                            r'(?:https?://|www\.)[^\s<>"\']+\bd[ée]sabonne(?:ment)?\b[^\s<>"\']*',
                            r'(?:https?://|www\.)[^\s<>"\']+\bannulation\b[^\s<>"\']*'
                        ]

                        for pattern in patterns:
                            matches = re.finditer(pattern, content, re.IGNORECASE)
                            for match in matches:
                                # Check if this URL is near keywords like "unsubscribe" or "désabonner"
                                unsubscribe_context = content[max(0, match.start() - 50):min(len(content), match.end() + 50)]
                                for keyword in ['unsubscribe', 'opt-out', 'opt out', 'désabonne', 'désabonnement', 'annulation']:
                                    if keyword in unsubscribe_context.lower():
                                        unsubscribe_url = match.group(0)
                                        break
                                if unsubscribe_url:
                                    break
                            if unsubscribe_url:
                                break

                except Exception as e:
                    pass  # Silently continue on content decoding errors

        # If it's a commercial email, add it to our collection
        if is_commercial and from_email:
            self.commercial_emails_found += 1
            domain = from_email.split('@')[-1]
            self.commercial_emails[domain].append({
                'from_name': from_name,
                'from_email': from_email,
                'subject': subject,
                'date': date.strftime("%Y-%m-%d %H:%M:%S") if date else 'Unknown',
                'unsubscribe_url': unsubscribe_url
            })

            # Store the most recent unsubscribe URL for this domain
            if unsubscribe_url and (domain not in self.unsubscribe_urls or date):
                if domain not in self.unsubscribe_urls or self.unsubscribe_urls[domain]['date'] is None or date > self.unsubscribe_urls[domain]['date']:
                    self.unsubscribe_urls[domain] = {
                        'url': unsubscribe_url,
                        'date': date
                    }

        # Get recipient email(s)
        for header in ['To', 'Delivered-To', 'X-Original-To', 'Envelope-To']:
            if msg.get(header):
                _, recipient_email = parseaddr(msg.get(header, ''))
                if recipient_email and '@' in recipient_email:
                    self.recipient_emails[recipient_email] += 1

    def generate_deletion_report(self, output_file='commercial_emails_report.csv'):
        """Generate a CSV report of commercial emails for deletion."""
        self.update_status(f"Generating CSV report: {output_file}")
        with open(output_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(['Domain', 'Email Count', 'First Email Date', 'Last Email Date',
                             'Sender Names', 'Sample Email', 'Unsubscribe URL'])

            for domain, emails in sorted(self.commercial_emails.items(),
                                         key=lambda x: len(x[1]), reverse=True):
                if emails:
                    # Sort emails by date
                    sorted_emails = sorted(emails, key=lambda x: x['date'])
                    first_date = sorted_emails[0]['date'] if sorted_emails else 'Unknown'
                    last_date = sorted_emails[-1]['date'] if sorted_emails else 'Unknown'

                    # Get unique sender names
                    sender_names = set(email['from_name'] for email in emails if email['from_name'])
                    sender_names_str = ', '.join(list(sender_names)[:3])
                    if len(sender_names) > 3:
                        sender_names_str += f", and {len(sender_names) - 3} more"

                    # Get sample email for reference
                    sample_email = sorted_emails[-1]['subject'] if sorted_emails else 'N/A'

                    # Get unsubscribe URL
                    unsubscribe_url = self.unsubscribe_urls.get(domain, {}).get('url', '')

                    writer.writerow([
                        domain,
                        len(emails),
                        first_date,
                        last_date,
                        sender_names_str,
                        sample_email,
                        unsubscribe_url
                    ])

        self.update_status(f"Report generated: {output_file}")
        return output_file

    def generate_deletion_instructions(self, output_file, language='en', user_name='', user_email=''):
        """
        Generate instructions for deleting commercial emails.
        
        Args:
            output_file (str): Path to the output file
            language (str): Language for instructions ('en', 'fr', or 'both')
            user_name (str): User's name to include in templates
            user_email (str): User's email to include in templates
        """
        # Instructions templates in different languages
        templates = {
            'en': {
                'header': "INSTRUCTIONS FOR DELETING COMMERCIAL EMAILS\n\n",
                'intro': "The following instructions will help you delete commercial emails from your account.\n\n",
                'step1': "1. For each domain below, search for emails from that domain in your email client.\n",
                'step2': "2. Review the emails and delete those you no longer need.\n",
                'step3': "3. Consider unsubscribing from mailing lists you no longer want to receive.\n\n",
                'domain_header': "DOMAIN: {domain} ({count} emails)\n",
                'search_query': "Search query: from:{domain}\n",
                'unsubscribe': "Unsubscribe request template:\n\n"
                               "Subject: Unsubscribe request and GDPR data deletion\n\n"
                               "Dear Sir/Madam,\n\n"
                               "I am writing to request the following actions regarding my personal data:\n\n"
                               "1. Please unsubscribe {email} from all your mailing lists and marketing communications.\n"
                               "2. As per my rights under the General Data Protection Regulation (GDPR), please delete all personal data you hold about me from your systems, including but not limited to my contact information, browsing history, purchase history, and any derived data profiles.\n"
                               "3. Please confirm in writing when this deletion has been completed.\n\n"
                               "If I do not receive confirmation within 30 days, I reserve the right to lodge a complaint with the relevant data protection authority.\n\n"
                               "Thank you for your prompt attention to this matter.\n\n"
                               "Regards,\n"
                               "{name}\n\n",
                'footer': "End of instructions. Good luck cleaning up your inbox!"
            },
            'fr': {
                'header': "INSTRUCTIONS POUR SUPPRIMER LES EMAILS COMMERCIAUX\n\n",
                'intro': "Les instructions suivantes vous aideront à supprimer les emails commerciaux de votre compte.\n\n",
                'step1': "1. Pour chaque domaine ci-dessous, recherchez les emails de ce domaine dans votre client de messagerie.\n",
                'step2': "2. Examinez les emails et supprimez ceux dont vous n'avez plus besoin.\n",
                'step3': "3. Envisagez de vous désabonner des listes de diffusion que vous ne souhaitez plus recevoir.\n\n",
                'domain_header': "DOMAINE: {domain} ({count} emails)\n",
                'search_query': "Requête de recherche: from:{domain}\n",
                'unsubscribe': "Modèle de demande de désabonnement:\n\n"
                               "Objet: Demande de désabonnement et suppression de données RGPD\n\n"
                               "Madame, Monsieur,\n\n"
                               "Je vous écris pour demander les actions suivantes concernant mes données personnelles :\n\n"
                               "1. Veuillez désabonner {email} de toutes vos listes de diffusion et communications marketing.\n"
                               "2. Conformément à mes droits en vertu du Règlement Général sur la Protection des Données (RGPD), veuillez supprimer toutes les données personnelles que vous détenez à mon sujet de vos systèmes, y compris, mais sans s'y limiter, mes coordonnées, mon historique de navigation, mon historique d'achat, et tous les profils de données dérivés.\n"
                               "3. Veuillez confirmer par écrit lorsque cette suppression aura été effectuée.\n\n"
                               "Si je ne reçois pas de confirmation dans les 30 jours, je me réserve le droit de déposer une plainte auprès de l'autorité de protection des données compétente.\n\n"
                               "Je vous remercie de votre attention concernant ma demande.\n\n"
                               "Cordialement,\n"
                               "{name}\n\n",
                'footer': "Fin des instructions. Bonne chance pour le nettoyage de votre boîte de réception!"
            }
        }

        # Determine the most common recipient email
        most_common_email = None
        if self.recipient_emails:
            most_common_email = self.recipient_emails.most_common(1)[0][0]
        
        # Use provided email, most common detected email, or placeholder
        email_placeholder = user_email if user_email else (most_common_email if most_common_email else "your.email@example.com")
        
        # Use user's name or a placeholder
        name_placeholder = user_name if user_name else "Your Name"

        with open(output_file, 'w', encoding='utf-8') as f:
            # Write header and intro based on language
            if language == 'en':
                f.write(templates['en']['header'])
                f.write(templates['en']['intro'])
                f.write(templates['en']['step1'])
                f.write(templates['en']['step2'])
                f.write(templates['en']['step3'])
            elif language == 'fr':
                f.write(templates['fr']['header'])
                f.write(templates['fr']['intro'])
                f.write(templates['fr']['step1'])
                f.write(templates['fr']['step2'])
                f.write(templates['fr']['step3'])
            else:  # both
                f.write(templates['en']['header'])
                f.write(templates['en']['intro'])
                f.write(templates['en']['step1'])
                f.write(templates['en']['step2'])
                f.write(templates['en']['step3'])
                f.write("\n\n---\n\n")
                f.write(templates['fr']['header'])
                f.write(templates['fr']['intro'])
                f.write(templates['fr']['step1'])
                f.write(templates['fr']['step2'])
                f.write(templates['fr']['step3'])

            # Write domain-specific instructions
            sorted_domains = sorted(self.commercial_emails.keys(),
                                   key=lambda x: len(self.commercial_emails[x]),
                                   reverse=True)

            for domain in sorted_domains:
                emails = self.commercial_emails[domain]
                count = len(emails)
                
                # Find the most common sender email for this domain
                sender_emails = [email['from_email'] for email in emails if email['from_email']]
                contact_email = domain  # Default fallback
                
                if sender_emails:
                    # Use the most recent sender email
                    contact_email = sender_emails[-1]
                    
                    # If it's just the domain, try to find a better email
                    if contact_email == domain:
                        for email in sender_emails:
                            if '@' in email and email.endswith(domain):
                                contact_email = email
                                break

                f.write("\n" + "-" * 50 + "\n\n")

                if language == 'en':
                    f.write(templates['en']['domain_header'].format(domain=domain, count=count))
                    f.write(templates['en']['search_query'].format(domain=domain))
                    f.write(templates['en']['unsubscribe'].format(email=email_placeholder, name=name_placeholder))
                elif language == 'fr':
                    f.write(templates['fr']['domain_header'].format(domain=domain, count=count))
                    f.write(templates['fr']['search_query'].format(domain=domain))
                    f.write(templates['fr']['unsubscribe'].format(email=email_placeholder, name=name_placeholder))
                else:  # both
                    f.write(templates['en']['domain_header'].format(domain=domain, count=count))
                    f.write(templates['en']['search_query'].format(domain=domain))
                    f.write("To: {}\n\n".format(contact_email))  # Add the contact email
                    f.write(templates['en']['unsubscribe'].format(email=email_placeholder, name=name_placeholder))
                    f.write("\n---\n\n")
                    f.write(templates['fr']['domain_header'].format(domain=domain, count=count))
                    f.write(templates['fr']['search_query'].format(domain=domain))
                    f.write("À: {}\n\n".format(contact_email))  # Add the contact email
                    f.write(templates['fr']['unsubscribe'].format(email=email_placeholder, name=name_placeholder))

            # Write footer
            f.write("\n" + "-" * 50 + "\n\n")
            if language == 'en':
                f.write(templates['en']['footer'])
            elif language == 'fr':
                f.write(templates['fr']['footer'])
            else:  # both
                f.write(templates['en']['footer'])
                f.write("\n\n---\n\n")
                f.write(templates['fr']['footer'])

    def analyze_large_attachments(self, export_dir, output_dir):
        """Analyze emails with large attachments and create a template for responses."""
        large_attachments = {}  # {sender_email: [(date, subject, size)]}
        size_threshold = 1048576  # 1024 * 1024  # 1MB in bytes

        # Walk through the export directory
        for root, _, files in os.walk(export_dir):
            for file in files:
                if file.endswith('.eml'):
                    email_path = os.path.join(root, file)
                    try:
                        with open(email_path, 'rb') as f:
                            msg = email.message_from_bytes(f.read())

                            # Get sender
                            from_header = msg.get('from', '')
                            sender_email = extract_email(from_header)
                            if not sender_email:
                                continue

                            # Get date and subject
                            date_str = msg.get('date', '')
                            try:
                                date = parsedate_to_datetime(date_str)
                            except:
                                date = None

                            subject = msg.get('subject', '(No subject)')

                            # Calculate total attachment size
                            total_size = 0
                            for part in msg.walk():
                                if part.get_content_maintype() == 'multipart':
                                    continue
                                if part.get('Content-Disposition') is None:
                                    continue

                                # Get attachment size
                                filename = part.get_filename()
                                if filename:
                                    payload = part.get_payload(decode=True)
                                    if payload:
                                        total_size += len(payload)

                            if total_size > size_threshold:
                                if sender_email not in large_attachments:
                                    large_attachments[sender_email] = []
                                large_attachments[sender_email].append((date, subject, total_size))

                    except Exception as e:
                        self.update_status(f"Error processing {email_path}: {str(e)}")

        # Generate the report
        if large_attachments:
            report_path = os.path.join(output_dir, 'large_attachments_report.txt')
            self.generate_large_attachments_report(large_attachments, report_path)
            return report_path
        return None

    def generate_large_attachments_report(self, large_attachments, output_file):
        """Generate a report for emails with large attachments."""
        templates = {
            'en': {
                'header': "LARGE ATTACHMENTS REPORT\n\n"
                         "The following users have sent emails with attachments larger than 1MB.\n"
                         "You may want to send them a friendly reminder about using file sharing services.\n\n",
                'email_template': "Email Template:\n\n"
                                "Subject: Regarding Large Email Attachments\n\n"
                                "Hello,\n\n"
                                "I noticed that you've sent me some emails with large attachments. "
                                "While I appreciate you sharing these files, large attachments can quickly fill up "
                                "email storage space and may be blocked by email servers.\n\n"
                                "For future reference, I recommend using file sharing services like:\n"
                                "- SwissTransfer (https://www.swisstransfer.com/)\n"
                                "- WeTransfer (https://wetransfer.com/)\n"
                                "- Google Drive (https://drive.google.com/)\n\n"
                                "These services are free, secure, and much more efficient for sharing large files.\n\n"
                                "Thank you for your understanding!\n\n"
                                "Best regards,\n"
                                "{name}\n\n",
                'sender_header': "\nSender: {}\n",
                'email_details': "  - Date: {}\n    Subject: {}\n    Size: {:.1f} MB\n"
            },
            'fr': {
                'header': "RAPPORT DES PIÈCES JOINTES VOLUMINEUSES\n\n"
                         "Les utilisateurs suivants ont envoyé des emails avec des pièces jointes supérieures à 1MB.\n"
                         "Vous pouvez leur envoyer un rappel amical concernant l'utilisation des services de partage de fichiers.\n\n",
                'email_template': "Modèle d'email:\n\n"
                                "Objet: Concernant les pièces jointes volumineuses\n\n"
                                "Bonjour,\n\n"
                                "J'ai remarqué que vous m'avez envoyé des emails avec des pièces jointes volumineuses. "
                                "Bien que j'apprécie que vous partagiez ces fichiers, les pièces jointes volumineuses peuvent "
                                "rapidement remplir l'espace de stockage des emails et peuvent être bloquées par les serveurs de messagerie.\n\n"
                                "Pour référence future, je recommande d'utiliser des services de partage de fichiers comme:\n"
                                "- SwissTransfer (https://www.swisstransfer.com/)\n"
                                "- WeTransfer (https://wetransfer.com/)\n"
                                "- Google Drive (https://drive.google.com/)\n\n"
                                "Ces services sont gratuits, sécurisés et beaucoup plus efficaces pour partager des fichiers volumineux.\n\n"
                                "Merci de votre compréhension!\n\n"
                                "Cordialement,\n"
                                "{name}\n\n",
                'sender_header': "\nExpéditeur: {}\n",
                'email_details': "  - Date: {}\n    Objet: {}\n    Taille: {:.1f} MB\n"
            }
        }

        with open(output_file, 'w', encoding='utf-8') as f:
            # Write header
            f.write(templates[self.current_language]['header'])

            # Write email template
            f.write(templates[self.current_language]['email_template'].format(
                name=self.name_var.get().strip() or "Your Name"
            ))

            f.write("-" * 50 + "\n")
            f.write("\nDetailed Report:\n")

            # Sort senders by total attachment size
            sorted_senders = sorted(
                large_attachments.items(),
                key=lambda x: sum(size for _, _, size in x[1]),
                reverse=True
            )

            # Write details for each sender
            for sender, attachments in sorted_senders:
                f.write(templates[self.current_language]['sender_header'].format(sender))

                # Sort attachments by date
                sorted_attachments = sorted(
                    attachments,
                    key=lambda x: x[0] if x[0] else datetime.min
                )

                for date, subject, size in sorted_attachments:
                    date_str = date.strftime("%Y-%m-%d %H:%M") if date else "Unknown date"
                    f.write(templates[self.current_language]['email_details'].format(
                        date_str, subject, size / (1024 * 1024)
                    ))

                # Add total for this sender
                total_size = sum(size for _, _, size in attachments)
                f.write(f"  Total: {total_size / (1024 * 1024):.1f} MB\n")


class AnalysisWorker(QThread):
    """Worker thread for email analysis."""
    status_signal = pyqtSignal(str)
    progress_signal = pyqtSignal(int, int)
    finished_signal = pyqtSignal(bool, str)  # Success flag and output directory

    def __init__(self, parser, export_dir, output_dir, language, user_name='', user_email=''):
        super().__init__()
        self.parser = parser
        self.export_dir = export_dir
        self.output_dir = output_dir
        self.language = language
        self.user_name = user_name
        self.user_email = user_email
        self.cancelled = False

    def run(self):
        try:
            # Create a new parser instance for the thread
            thread_parser = EmailExportParser()
            thread_parser.set_status_callback(lambda msg: self.status_signal.emit(msg))
            thread_parser.set_progress_callback(lambda curr, total: self.progress_signal.emit(curr, total))
            thread_parser.current_language = self.language  # Set the language
            
            thread_parser.export_dir = self.export_dir
            thread_parser.parse_export()

            if self.cancelled:
                self.finished_signal.emit(False, "")
                return

            # Generate commercial emails report
            csv_path = os.path.join(self.output_dir, 'commercial_emails_report.csv')
            thread_parser.generate_deletion_report(csv_path)

            # Generate deletion instructions
            instructions_path = os.path.join(self.output_dir, 'deletion_instructions.txt')
            thread_parser.generate_deletion_instructions(
                output_file=instructions_path,
                language=self.language,
                user_name=self.user_name,
                user_email=self.user_email
            )

            # Analyze large attachments
            self.status_signal.emit("\nAnalyzing emails with large attachments...")
            large_attachments_report = thread_parser.analyze_large_attachments(
                self.export_dir,
                self.output_dir
            )

            # Emit summary
            self.status_signal.emit(f"\nAnalysis complete!")
            self.status_signal.emit(f"Found {sum(len(emails) for emails in thread_parser.commercial_emails.values())} commercial emails")
            self.status_signal.emit(f"From {len(thread_parser.commercial_emails)} unique domains")
            self.status_signal.emit(f"\nReports saved to:")
            self.status_signal.emit(f"- {csv_path}")
            self.status_signal.emit(f"- {instructions_path}")
            if large_attachments_report:
                self.status_signal.emit(f"- {large_attachments_report}")

            self.finished_signal.emit(True, self.output_dir)

        except Exception as e:
            self.status_signal.emit(f"Error: {str(e)}")
            self.finished_signal.emit(False, "")


class EmailParserCLI:
    """Command-line interface for the Email Export Parser."""

    def __init__(self):
        self.parser = EmailExportParser()
        self.parser.set_status_callback(self.update_status)

    def update_status(self, message):
        """Print status updates to the console."""
        print(message)

    def run(self):
        """Run the command-line interface."""
        # Parse command-line arguments
        arg_parser = argparse.ArgumentParser(
            description='Parse email exports from Gmail and Outlook to identify commercial senders'
        )
        arg_parser.add_argument('export_dir', help='Directory containing email export files')
        arg_parser.add_argument(
            '--output-dir',
            help='Directory to save output reports (defaults to export directory)',
            default=None
        )
        arg_parser.add_argument(
            '--language',
            help='Language for deletion instructions: en, fr, or both',
            choices=['en', 'fr', 'both'],
            default='en'
        )

        args = arg_parser.parse_args()

        # Set export directory
        self.parser.export_dir = args.export_dir

        # Set output directory
        output_dir = args.output_dir if args.output_dir else args.export_dir
        os.makedirs(output_dir, exist_ok=True)

        try:
            # Parse export
            print(f"Starting email export analysis from: {args.export_dir}")
            self.parser.parse_export()

            # Generate reports
            csv_path = os.path.join(output_dir, 'commercial_emails_report.csv')
            self.parser.generate_deletion_report(csv_path)

            instructions_path = os.path.join(output_dir, 'deletion_instructions.txt')
            self.parser.generate_deletion_instructions(
                output_file=instructions_path,
                language=args.language
            )

            print("\nAnalysis complete!")
            print(f"Found {sum(len(emails) for emails in self.parser.commercial_emails.values())} commercial emails")
            print(f"From {len(self.parser.commercial_emails)} unique domains")
            print(f"\nReports saved to:")
            print(f"- {csv_path}")
            print(f"- {instructions_path}")

        except Exception as e:
            print(f"Error: {str(e)}")
            traceback.print_exc()
            return 1

        return 0


class EmailParserGUI(QMainWindow):
    """PyQt6-based GUI for the Email Export Parser."""

    def __init__(self):
        super().__init__()
        self.parser = EmailExportParser()
        self.parser.set_status_callback(self.update_status)
        self.parser.set_progress_callback(self.update_progress)

        self.system_language = self.detect_system_language()
        self.current_language = self.system_language

        self.setup_ui()
        self.worker = None

    def detect_system_language(self):
        """Detect the system language."""
        try:
            import locale
            lang_code = locale.getdefaultlocale()[0]
            if lang_code and lang_code.startswith('fr'):
                return 'fr'
        except:
            pass
        return 'en'

    def setup_ui(self):
        """Set up the user interface."""
        self.setWindowTitle("Email Export Parser")
        self.setMinimumSize(800, 700)

        # Create central widget and main layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # Language selection
        self.lang_groupLab = QGroupBox(self.get_text('lang_settings'))
        lang_layout = QHBoxLayout()
        self.lang_group = QButtonGroup()

        self.en_radio = QRadioButton("English")
        self.fr_radio = QRadioButton("Français")
        self.lang_group.addButton(self.en_radio)
        self.lang_group.addButton(self.fr_radio)

        if self.current_language == 'fr':
            self.fr_radio.setChecked(True)
        else:
            self.en_radio.setChecked(True)

        self.lang_label = QLabel(self.get_text('language'))

        lang_layout.addWidget(self.lang_label)
        lang_layout.addWidget(self.en_radio)
        lang_layout.addWidget(self.fr_radio)
        lang_layout.addStretch()
        self.lang_groupLab.setLayout(lang_layout)
        main_layout.addWidget(self.lang_groupLab)

        # Directory selection
        dir_frame = QFrame()
        dir_layout = QGridLayout()

        self.dir_label = QLabel(self.get_text('export_dir'))
        self.dir_edit = QLineEdit()
        self.browse_btn = QPushButton(self.get_text('browse'))

        dir_layout.addWidget(self.dir_label, 0, 0)
        dir_layout.addWidget(self.dir_edit, 0, 1)
        dir_layout.addWidget(self.browse_btn, 0, 2)

        self.out_dir_label = QLabel(self.get_text('output_dir'))
        self.out_dir_edit = QLineEdit()
        self.browse_out_btn = QPushButton(self.get_text('browse'))

        dir_layout.addWidget(self.out_dir_label, 1, 0)
        dir_layout.addWidget(self.out_dir_edit, 1, 1)
        dir_layout.addWidget(self.browse_out_btn, 1, 2)

        dir_frame.setLayout(dir_layout)
        main_layout.addWidget(dir_frame)

        # User information
        self.user_group = QGroupBox(self.get_text('user_info'))
        user_layout = QGridLayout()

        self.name_label = QLabel(self.get_text('your_name'))
        self.name_edit = QLineEdit()
        self.email_label = QLabel(self.get_text('your_email'))
        self.email_edit = QLineEdit()

        user_layout.addWidget(self.name_label, 0, 0)
        user_layout.addWidget(self.name_edit, 0, 1)
        user_layout.addWidget(self.email_label, 1, 0)
        user_layout.addWidget(self.email_edit, 1, 1)

        self.user_group.setLayout(user_layout)
        main_layout.addWidget(self.user_group)

        # Action buttons
        btn_frame = QFrame()
        btn_layout = QHBoxLayout()

        self.analyze_btn = QPushButton(self.get_text('analyze'))
        self.cancel_btn = QPushButton(self.get_text('cancel'))
        self.cancel_btn.setEnabled(False)

        btn_layout.addWidget(self.analyze_btn)
        btn_layout.addWidget(self.cancel_btn)
        btn_layout.addStretch()

        btn_frame.setLayout(btn_layout)
        main_layout.addWidget(btn_frame)

        # Progress bar
        progress_frame = QFrame()
        progress_layout = QHBoxLayout()

        self.progress_label = QLabel(self.get_text('progress'))
        self.progress_bar = QProgressBar()

        progress_layout.addWidget(self.progress_label)
        progress_layout.addWidget(self.progress_bar)

        progress_frame.setLayout(progress_layout)
        main_layout.addWidget(progress_frame)

        # Status log
        self.status_text = QTextEdit()
        self.status_text.setReadOnly(True)
        main_layout.addWidget(self.status_text)

        # Connect signals
        self.browse_btn.clicked.connect(self.browse_directory)
        self.browse_out_btn.clicked.connect(self.browse_output_directory)
        self.analyze_btn.clicked.connect(self.start_analysis)
        self.cancel_btn.clicked.connect(self.cancel_analysis)
        self.en_radio.toggled.connect(self.update_language)
        self.fr_radio.toggled.connect(self.update_language)

        # Set initial status
        self.update_status(self.get_text('ready'))

        # Logo and credits frame
        logo_frame = QFrame()
        logo_layout = QVBoxLayout()

        # Top row with logo and company link
        top_row = QHBoxLayout()

        # Add clickable logo
        logo_label = QLabel()
        logo_path = os.path.join('assets', 'les-enovateurs-logo.webp')
        if os.path.exists(logo_path):
            logo_pixmap = QPixmap(logo_path)
            logo_label.setPixmap(logo_pixmap.scaled(200, 50, Qt.AspectRatioMode.KeepAspectRatio))
        logo_label.setCursor(Qt.CursorShape.PointingHandCursor)
        logo_label.mousePressEvent = lambda _: QDesktopServices.openUrl(QUrl("https://les-enovateurs.com"))

        # Add social links
        social_links = QHBoxLayout()
        social_links.addStretch()

        self.linkedin_label = QLabel(self.get_text('follow_us_linkedin'))
        self.linkedin_label.setOpenExternalLinks(True)

        self.mastodon_label = QLabel(self.get_text('follow_us_mastodon'))
        self.mastodon_label.setOpenExternalLinks(True)

        social_links.addWidget(self.linkedin_label)
        social_links.addWidget(self.mastodon_label)

        top_row.addWidget(logo_label)
        top_row.addStretch()
        top_row.addLayout(social_links)

        # Bottom row with creator credit
        bottom_row = QHBoxLayout()
        self.creator_label = QLabel(self.get_text('creator_credit'))
        self.creator_label.setOpenExternalLinks(True)
        self.creator_label.setAlignment(Qt.AlignmentFlag.AlignRight)
        bottom_row.addStretch()
        bottom_row.addWidget(self.creator_label)

        # Add rows to logo frame
        logo_layout.addLayout(top_row)
        logo_layout.addLayout(bottom_row)

        logo_frame.setLayout(logo_layout)
        main_layout.addWidget(logo_frame)

    def get_text(self, key):
        """Get UI text based on current language."""
        texts = {
            'en': {
                'window_title': "Email Export Parser",
                'export_dir': "Email Export Directory:",
                'output_dir': "Output Directory:",
                'browse': "Browse...",
                'user_info': "User Information (for deletion instructions)",
                'your_name': "Your Name:",
                'your_email': "Your Email:",
                'lang_settings': "Language Settings",
                'language': "Language:",
                'english': "English",
                'french': "French",
                'analyze': "Analyze Emails",
                'cancel': "Cancel",
                'progress': "Progress:",
                'status': "Status:",
                'ready': "Ready. Select an email export directory and click 'Analyze Emails'.",
                'analysis_complete': "Analysis complete! Open the output directory?",
                'error_select_dir': "Please select an email export directory.",
                'analysis_cancelled': "Analysis cancelled.",
                'analysis_started': "Starting analysis...",
                'found_emails': "Found {} commercial emails",
                'from_domains': "From {} unique domains",
                'reports_saved': "Reports saved to:",
                'error_occurred': "Error: {}",
                'yes': "Yes",
                'no': "No",
                'follow_us_linkedin': "<a href=\"https://www.linkedin.com/company/les-enovateurs\">Follow us on LinkedIn</a>",
                'follow_us_mastodon': "<a href=\"https://mastodon.social/@enovateurs_media\">Follow us on Mastodon</a>",
                'creator_credit': "Created by <a href=\"https://www.linkedin.com/in/jeremy-pastouret\">Jérémy Pastouret</a> for <a href=\"https://les-enovateurs.com\">Les e-novateurs</a>"
            },
            'fr': {
                'window_title': "Analyseur d'Exportation d'Emails",
                'export_dir': "Répertoire d'exportation d'emails:",
                'output_dir': "Répertoire de sortie:",
                'browse': "Parcourir...",
                'user_info': "Informations utilisateur (pour les instructions)",
                'your_name': "Votre nom:",
                'your_email': "Votre email:",
                'lang_settings': "Paramètres de langue",
                'language': "Langue:",
                'english': "Anglais",
                'french': "Français",
                'analyze': "Analyser les Emails",
                'cancel': "Annuler",
                'progress': "Progression:",
                'status': "Statut:",
                'ready': "Prêt. Sélectionnez un répertoire d'exportation d'emails et cliquez sur 'Analyser les Emails'.",
                'analysis_complete': "Analyse terminée ! Ouvrir le répertoire de sortie ?",
                'error_select_dir': "Veuillez sélectionner un répertoire d'exportation d'emails.",
                'analysis_cancelled': "Analyse annulée.",
                'analysis_started': "Démarrage de l'analyse...",
                'found_emails': "{} emails commerciaux trouvés",
                'from_domains': "Provenant de {} domaines uniques",
                'reports_saved': "Rapports sauvegardés dans:",
                'error_occurred': "Erreur: {}",
                'yes': "Oui",
                'no': "Non",
                'follow_us_linkedin': "<a href=\"https://www.linkedin.com/company/les-enovateurs\">Suivez-nous sur LinkedIn</a>",
                'follow_us_mastodon': "<a href=\"https://mastodon.social/@enovateurs_media\">Suivez-nous sur Mastodon</a>",
                'creator_credit': "Initié par <a href=\"https://www.linkedin.com/in/jeremy-pastouret\">Jérémy Pastouret</a> pour <a href=\"https://les-enovateurs.com\">Les e-novateurs</a>"
            }
        }
        return texts[self.current_language][key]

    def update_language(self):
        """Update the interface language."""
        # Update the current language based on radio button selection
        self.current_language = 'fr' if self.fr_radio.isChecked() else 'en'

        # Update window title
        self.setWindowTitle(self.get_text('window_title'))

        # Update all labels and buttons
        self.dir_label.setText(self.get_text('export_dir'))
        self.out_dir_label.setText(self.get_text('output_dir'))
        self.browse_btn.setText(self.get_text('browse'))
        self.browse_out_btn.setText(self.get_text('browse'))
        self.lang_label.setText(self.get_text('language'))
        self.lang_groupLab.setTitle(self.get_text('lang_settings'))

        # Update user info group box
        self.user_group.setTitle(self.get_text('user_info'))

        self.name_label.setText(self.get_text('your_name'))
        self.email_label.setText(self.get_text('your_email'))
        self.analyze_btn.setText(self.get_text('analyze'))
        self.cancel_btn.setText(self.get_text('cancel'))
        self.progress_label.setText(self.get_text('progress'))

        # Clear and update the status text
        self.status_text.clear()
        self.update_status(self.get_text('ready'))

        # Update the creator credit
        self.creator_label.setText(self.get_text('creator_credit'))

        # Update the social links
        self.linkedin_label.setText(self.get_text('follow_us_linkedin'))
        self.mastodon_label.setText(self.get_text('follow_us_mastodon'))

    def browse_directory(self):
        """Open a directory browser dialog."""
        directory = QFileDialog.getExistingDirectory(self, self.get_text('export_dir'))
        if directory:
            self.dir_edit.setText(directory)
            # Default output to same directory
            if not self.out_dir_edit.text():
                self.out_dir_edit.setText(directory)

    def browse_output_directory(self):
        """Open a directory browser dialog for output."""
        directory = QFileDialog.getExistingDirectory(self, self.get_text('output_dir'))
        if directory:
            self.out_dir_edit.setText(directory)

    def update_status(self, message):
        """Update the status text widget."""
        self.status_text.append(message)

    def update_progress(self, current, total=None):
        """Update the progress bar."""
        if total:
            progress_pct = int((current / total) * 100)
            self.progress_bar.setValue(progress_pct)
        else:
            # Indeterminate progress
            self.progress_bar.setValue(current % 100)

    def start_analysis(self):
        """Start the email analysis in a separate thread."""
        export_dir = self.dir_edit.text()
        output_dir = self.out_dir_edit.text()
        user_name = self.name_edit.text()
        user_email = self.email_edit.text()
        
        if not export_dir:
            QMessageBox.warning(self, self.get_text('window_title'), 
                               self.get_text('error_select_dir'))
            return
        
        if not output_dir:
            output_dir = export_dir
        
        # Create output directory if it doesn't exist
        os.makedirs(output_dir, exist_ok=True)
        
        # Reset progress and status
        self.progress_bar.setValue(0)
        self.status_text.clear()
        self.update_status(self.get_text('analysis_started'))
        
        # Update UI state
        self.analyze_btn.setEnabled(False)
        self.cancel_btn.setEnabled(True)
        
        # Create a new worker instance
        if self.worker:
            self.worker.cancelled = True
            self.worker.wait()
        
        self.worker = AnalysisWorker(
            self.parser, 
            export_dir, 
            output_dir, 
            self.current_language,
            user_name,  # Pass the user's name
            user_email  # Pass the user's email
        )
        self.worker.status_signal.connect(self.update_status)
        self.worker.progress_signal.connect(self.update_progress)
        self.worker.finished_signal.connect(self.finish_analysis)
        self.worker.start()

    def cancel_analysis(self):
        """Cancel the ongoing analysis."""
        if self.worker and self.worker.isRunning():
            self.worker.cancelled = True
            self.update_status(self.get_text('analysis_cancelled'))

    def finish_analysis(self, success, output_dir):
        """Reset UI after analysis is complete."""
        self.analyze_btn.setEnabled(True)
        self.cancel_btn.setEnabled(False)

        if success:
            self.update_status(self.get_text('analysis_complete'))
            self.ask_open_directory(output_dir)
        else:
            self.update_status(self.get_text('error_occurred').format(output_dir))

    def ask_open_directory(self, directory):
        """Ask if the user wants to open the output directory."""
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle(self.get_text('window_title'))
        msg_box.setText(self.get_text('analysis_complete'))
        msg_box.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        
        # Set translated button text
        yes_button = msg_box.button(QMessageBox.StandardButton.Yes)
        yes_button.setText(self.get_text('yes'))
        
        no_button = msg_box.button(QMessageBox.StandardButton.No)
        no_button.setText(self.get_text('no'))
        
        reply = msg_box.exec()
        
        if reply == QMessageBox.StandardButton.Yes:
            self.open_directory(directory)

    def open_directory(self, directory):
        """Open the specified directory in the file explorer."""
        try:
            if sys.platform == 'win32':
                os.startfile(directory)
            elif sys.platform == 'darwin':  # macOS
                import subprocess
                subprocess.call(['open', directory])
            else:  # Linux
                import subprocess
                subprocess.call(['xdg-open', directory])
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Could not open directory: {str(e)}")

    def run(self):
        """Run the GUI application."""
        self.show()


def main():
    """Main entry point for the application."""
    if len(sys.argv) > 1:
        # CLI mode
        print("\nEmail Export Parser")
        print("Powered by Les E-novateurs (https://les-enovateurs.com)")
        print("-" * 50 + "\n")
        cli = EmailParserCLI()
        sys.exit(cli.run())
    else:
        # GUI mode
        app = QApplication(sys.argv)
        gui = EmailParserGUI()
        gui.show()
        sys.exit(app.exec())


if __name__ == "__main__":
    main()