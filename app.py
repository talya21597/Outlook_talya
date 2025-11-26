# -*- coding: utf-8 -*-
from flask import Flask, render_template, request, jsonify
import win32com.client
import pythoncom
import os
from werkzeug.utils import secure_filename

app = Flask(__name__)

# הגדרות
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'pdf', 'doc', 'docx', 'txt'}

# וודא שתיקיית uploads קיימת
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # מקסימום 16MB


def allowed_file(filename):
    """בדיקה האם סיומת הקובץ מותרת"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def create_outlook_draft(subject, recipient, body, attachment_path=None):
    """
    פונקציה שיוצרת טיוטת מייל ב-Outlook
    
    Args:
        subject: נושא המייל
        recipient: כתובת הנמען
        body: גוף ההודעה
        attachment_path: נתיב לקובץ מצורף (אופציונלי)
    """
    try:
        # אתחול COM - חובה ל-Windows threading!
        pythoncom.CoInitialize()
        
        # התחברות ל-Outlook (יפתח אוטומטית אם סגור!)
        outlook = win32com.client.Dispatch("Outlook.Application")
        
        # קבלת namespace (נדרש כדי לוודא ש-Outlook מוכן)
        namespace = outlook.GetNamespace("MAPI")
        
        # יצירת מייל חדש (0 = olMailItem)
        mail = outlook.CreateItem(0)
        
        # מילוי פרטי המייל
        mail.Subject = subject
        mail.To = recipient
        
        # שימוש ב-HTMLBody כדי לוודא קידוד נכון של עברית
        # המרת שורות חדשות ל-HTML
        html_body = body.replace('\r\n', '<br>').replace('\n', '<br>')
        mail.HTMLBody = f'<html><head><meta charset="UTF-8"></head><body style="font-family: Arial; font-size: 11pt;">{html_body}</body></html>'
        
        # צירוף קובץ אם קיים
        if attachment_path and os.path.exists(attachment_path):
            mail.Attachments.Add(os.path.abspath(attachment_path))
        
        # שמירת הטיוטה (לא הצגה מיידית!)
        mail.Save()
        
        # הצגת הטיוטה ללא חסימה
        mail.Display(False)
        
        return True
    
    except Exception as e:
        print(f"שגיאה ביצירת טיוטה עבור {recipient}: {str(e)}")
        import traceback
        traceback.print_exc()
        return False
    
    finally:
        # ניקוי COM
        pythoncom.CoUninitialize()


@app.route('/')
def index():
    """עמוד הבית - טעינת הטופס"""
    return render_template('index.html')


@app.route('/send', methods=['POST'])
def send_emails():
    """
    מעבד את הטופס ויוצר טיוטות ב-Outlook
    """
    try:
        # קבלת נתונים מהטופס
        subject = request.form.get('subject', '').strip()
        recipients = request.form.get('recipients', '').strip()
        body = request.form.get('body', '').strip()
        
        # בדיקת שדות חובה
        if not subject or not recipients or not body:
            return jsonify({
                'success': False,
                'message': 'יש למלא את כל השדות'
            }), 400
        
        # פיצול הנמענים (מופרדים בפסיקים או נקודה-פסיק)
        recipients_list = [r.strip() for r in recipients.replace(';', ',').split(',') if r.strip()]
        
        if not recipients_list:
            return jsonify({
                'success': False,
                'message': 'לא הוזנו כתובות נמענים תקינות'
            }), 400
        
        # טיפול בקובץ מצורף
        attachment_path = None
        if 'attachment' in request.files:
            file = request.files['attachment']
            if file and file.filename and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                attachment_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                file.save(attachment_path)
        
        # יצירת טיוטה עבור כל נמען
        success_count = 0
        failed_recipients = []
        
        for recipient in recipients_list:
            if create_outlook_draft(subject, recipient, body, attachment_path):
                success_count += 1
            else:
                failed_recipients.append(recipient)
        
        # הכנת תשובה
        if success_count == len(recipients_list):
            message = f'נפתחו בהצלחה {success_count} טיוטות ב-Outlook!'
            return jsonify({
                'success': True,
                'message': message,
                'count': success_count
            })
        elif success_count > 0:
            message = f'נפתחו {success_count} טיוטות. נכשלו: {", ".join(failed_recipients)}'
            return jsonify({
                'success': True,
                'message': message,
                'count': success_count,
                'failed': failed_recipients
            })
        else:
            return jsonify({
                'success': False,
                'message': 'לא הצלחנו לפתוח אף טיוטה. בדקי ש-Outlook פתוח.'
            }), 500
    
    except Exception as e:
        return jsonify({
            'success': False,
            'message': f'שגיאה: {str(e)}'
        }), 500


if __name__ == '__main__':
    print("🚀 השרת רץ על: http://localhost:5000")
    print("📧 פתחי את הדפדפן וגלשי לכתובת למעלה")
    print("⚠️  וודאי ש-Outlook פתוח!")
    app.run(debug=True, port=5000)