import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph

def load_data(filename):
    return pd.read_excel(filename, engine='openpyxl')

def generate_report_cards(df):
    subjects = ['Science', 'English', 'History', 'Maths']
    
    for _, row in df.iterrows():
        student_id = row['id']
        name = row['Name']
        total_score = sum(row[subject] for subject in subjects)
        average_score = total_score / len(subjects)
        
        pdf_filename = f"report_card_{student_id}.pdf"
        doc = SimpleDocTemplate(pdf_filename, pagesize=A4)
        
        styles = getSampleStyleSheet()
        title = Paragraph(f"Report Card for {name} (ID: {student_id})", styles['Title'])
        summary = Paragraph(f"Total Score: {total_score}<br/>Average Score: {average_score:.2f}", styles['Normal'])
        
        data = [['Subject', 'Score']] + [[subject, row[subject]] for subject in subjects]
        table = Table(data)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige)
        ]))
        
        doc.build([title, summary, table])
        print(f"Generated {pdf_filename}")

if __name__ == "__main__":
    filename = "/content/marksheet.xlsx"  # Ensure this is the correct path to your Excel file
    df = load_data(filename)
    generate_report_cards(df)
