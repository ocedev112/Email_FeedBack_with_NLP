import nltk
nltk.download('punkt')
nltk.download('averaged_perceptron_tagger')
from dotenv import load_dotenv
import imaplib
import email
from email.header import decode_header
import pandas as pd
from textblob import TextBlob
from datetime import datetime, timedelta
import os

load_dotenv()



EMAIL = os.getenv('EMAIL')
PASSWORD = os.getenv('PASSWORD')
IMAP_SERVER = 'imap.gmail.com'
MAILBOX = 'INBOX'
SUBJECT_KEYWORDS=['feedback','review','opinion']

def connect_to_email_and_search():
    mail = imaplib.IMAP4_SSL(IMAP_SERVER)
    mail.login(EMAIL,PASSWORD)
    mail.select(MAILBOX)
    date_since = (datetime.now() - timedelta(days=7)).strftime("%d-%b-%Y")
    result, data = mail.search(None, f'(SINCE "{date_since}")')
    return mail, data[0].split()




def extract_feedback(mail,email_ids):
    feedbacks = []

    for eid in email_ids:
        res, msg_data = mail.fetch(eid, '(RFC822)')
        for response_part in msg_data:
            if isinstance(response_part,tuple):
                msg  = email.message_from_bytes(response_part[1])
                subject = decode_header(msg["Subject"])[0][0]
                if isinstance(subject,bytes):
                    subject = subject.decode()

                if any(keyword.lower() in subject.lower() for keyword in SUBJECT_KEYWORDS):
                      body = "" 
                      if msg.is_multipart():
                          for part in msg.walk():
                              if part.get_content_type() == "text/plain":
                                  body = part.get_payload(decode=True)
                                  break
                      else:
                            body = msg.get_payload(decode=True)
                    
                      feedbacks.append({
                         'date':msg["Date"],
                         'from':msg["From"],
                         'subject':subject,
                         'body': body.strip()
                       })
    return feedbacks


def analyze_sentiment(feedbacks):
    for fb in feedbacks:
        analysis = TextBlob(fb['body'])
        fb['polarity'] = analysis.sentiment.polarity
        fb['sentiment'] =  "Positive" if analysis.polarity >0 else "Negative" if analysis.polarity < 0 else 'Neutral'
    return feedbacks

def generate_report(feedbacks):
    df = pd.DataFrame(feedbacks)
    filename = f"weekly_feedback_report_{datetime.now().strftime('%Y%m%d')}.csv"
    df.to_csv(filename, index=False)
    print(f" Report saved as {filename}")


def main():
    print("Connecting to email....")
    mail, email_ids = connect_to_email_and_search()
    feedbacks =extract_feedback(mail,email_ids)

    if not feedbacks:
        print("No feedback emails found")
    
    print("Analyzing Sentiment")
    feedbacks = analyze_sentiment(feedbacks)

    generate_report(feedbacks)

if __name__ == "__main__":
    main()