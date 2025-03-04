import streamlit as st
from docx import Document
import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
import io
import re
import tempfile
import os

st.set_page_config(page_title="ממיר הערות וורד לאקסל", page_icon="📝")

st.title("ממיר הערות מקובץ וורד לאקסל")
st.markdown("אפליקציה זו ממירה את כל ההערות והתגובות מקובץ וורד לקובץ אקסל מסודר.")

def extract_comments_from_docx(docx_file):
    """
    מחלץ הערות ותגובות מקובץ וורד
    
    :param docx_file: קובץ וורד (כאובייקט BytesIO)
    :return: רשימה של הערות ותגובות
    """
    # מידע שנרצה לשמור לגבי כל הערה
    comments_data = []
    
    # שמירת הקובץ הזמני
    with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_file:
        tmp_file.write(docx_file.getvalue())
        tmp_path = tmp_file.name
    
    try:
        # קובץ וורד הוא למעשה קובץ ZIP שמכיל מסמכי XML
        with zipfile.ZipFile(tmp_path, 'r') as zip_ref:
            # בדיקה אם יש קובץ הערות
            comment_files = [f for f in zip_ref.namelist() if 'comments.xml' in f]
            
            if not comment_files:
                return comments_data
            
            # קריאת מסמך הוורד לצורך קבלת הטקסט המקורי
            doc = Document(tmp_path)
            paragraphs = [p.text for p in doc.paragraphs]
            
            # מיפוי מספרי פסקאות לעמודים (הערכה פשוטה)
            # זה אינו מדויק ב-100% ודורש ספריה חיצונית לקבלת מספרי עמודים מדויקים
            para_to_page = {}
            approx_chars_per_page = 3000  # הערכה של כמות תווים בעמוד
            current_page = 1
            char_count = 0
            
            for i, para in enumerate(paragraphs):
                char_count += len(para)
                para_to_page[i] = current_page
                if char_count > approx_chars_per_page:
                    current_page += 1
                    char_count = 0
            
            # קריאת קובץ XML של ההערות
            for comment_file in comment_files:
                xml_content = zip_ref.read(comment_file)
                root = ET.fromstring(xml_content)
                
                # מציאת המרחב שמות (namespace)
                ns = {'w': re.search(r'{(.*)}', root.tag).group(1)}
                
                # חילוץ כל ההערות
                comments = root.findall('.//w:comment', ns)
                
                for comment in comments:
                    comment_id = comment.get(f"{{{ns['w']}}}id")
                    author = comment.get(f"{{{ns['w']}}}author", "לא ידוע")
                    date = comment.get(f"{{{ns['w']}}}date", "")
                    
                    # טקסט ההערה
                    comment_text_elements = comment.findall('.//w:t', ns)
                    comment_text = "".join([elem.text for elem in comment_text_elements if elem.text])
                    
                    # חיפוש תגובות (הערות מקושרות)
                    parent_id = comment.get(f"{{{ns['w']}}}parentId")
                    
                    # הוספת מידע ההערה לרשימה
                    comment_data = {
                        'id': comment_id,
                        'parent_id': parent_id,
                        'author': author,
                        'date': date,
                        'text': comment_text,
                        'page': 0  # ערך ברירת מחדל, יעודכן בהמשך
                    }
                    
                    comments_data.append(comment_data)
        
        # יצירת מבנה נתונים היררכי של הערות ותגובות
        comments_dict = {}
        for comment in comments_data:
            comments_dict[comment['id']] = comment
        
        # מסדר הערות ותגובות
        structured_comments = []
        
        # מיון ההערות לפי עיקריות ותגובות
        for comment in comments_data:
            if comment['parent_id'] is None:  # הערה עיקרית
                # חיפוש כל התגובות להערה זו
                replies = []
                for reply in comments_data:
                    if reply['parent_id'] == comment['id']:
                        replies.append(reply)
                
                # יצירת שורה חדשה לאקסל
                row = {
                    'הערה': comment['text'],
                    'כותב': comment['author'],
                    'עמוד': para_to_page.get(0, 1),  # כברירת מחדל עמוד 1
                    'תאריך': comment['date']
                }
                
                # הוספת תגובות
                for i, reply in enumerate(replies, 1):
                    row[f'תגובה {i}'] = reply['text']
                    row[f'כותב תגובה {i}'] = reply['author']
                    row[f'תאריך תגובה {i}'] = reply['date']
                
                structured_comments.append(row)
        
        return structured_comments
    finally:
        # מחיקת הקובץ הזמני
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)

# טעינת קובץ
uploaded_file = st.file_uploader("העלה קובץ וורד (.docx)", type=['docx'])

if uploaded_file is not None:
    with st.spinner('מעבד את הקובץ...'):
        # חילוץ ההערות והתגובות
        comments_data = extract_comments_from_docx(uploaded_file)
        
        if not comments_data:
            st.warning("לא נמצאו הערות בקובץ.")
        else:
            # יצירת DataFrame
            df = pd.DataFrame(comments_data)
            
            # הצגת תצוגה מקדימה
            st.subheader("תצוגה מקדימה של ההערות")
            st.dataframe(df)
            
            # יצירת קובץ אקסל בזיכרון
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='הערות')
            
            # כפתור להורדת הקובץ
            excel_data = output.getvalue()
            file_name = uploaded_file.name.replace('.docx', '_comments.xlsx')
            
            st.download_button(
                label="הורד קובץ אקסל",
                data=excel_data,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            # מידע סטטיסטי
            st.subheader("סיכום")
            st.write(f"סה״כ נמצאו: {len(df)} הערות עיקריות")
            
            # בדיקה כמה הערות יש להן תגובות
            replies_columns = [col for col in df.columns if 'תגובה ' in col and 'כותב' not in col and 'תאריך' not in col]
            if replies_columns:
                has_replies = df[replies_columns[0]].notna().sum()
                st.write(f"מתוכן {has_replies} הערות עם תגובות")

# הוספת הוראות שימוש
with st.expander("הוראות שימוש"):
    st.markdown("""
    ### איך להשתמש באפליקציה:
    1. לחץ על "העלה קובץ וורד" ובחר את הקובץ שלך (בפורמט .docx)
    2. האפליקציה תעבד את הקובץ ותחלץ את כל ההערות והתגובות
    3. תוצג תצוגה מקדימה של הנתונים שחולצו
    4. לחץ על "הורד קובץ אקסל" כדי לשמור את הנתונים כקובץ אקסל
    
    ### הערות:
    * מספרי העמודים הם הערכה בלבד ועשויים להיות לא מדויקים במסמכים מורכבים
    * הקובץ שלך לא נשמר בשרת ומעובד באופן מקומי בלבד
    """)

# פוטר
st.markdown("---")
st.markdown("פותח עם ❤️ באמצעות Streamlit")
