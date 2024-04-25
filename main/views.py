from io import BytesIO
from django.shortcuts import render
from django.http import HttpResponse 
import os
import tempfile
import fitz
import re
import docx2txt
import xlwt

def main(request):
    return render(request, 'index.html')

def extract_text_from_pdf(pdf_path):
    text = ''
    with fitz.open(pdf_path) as doc:
        for page in doc:
            text += page.get_text()
    return text

def extract_text_from_docx(docx_path):
    text = docx2txt.process(docx_path)
    return text

def extract_info_from_text(text):
    email_pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
    phone_pattern = r'\b(?:\d{3}[-.\s]?)?\d{3}[-.\s]?\d{4}\b'
    name_pattern = r'\b[A-Z][a-z]+ [A-Z][a-z]+\b'

    emails = re.findall(email_pattern, text)
    phones = re.findall(phone_pattern, text)
    tags = re.findall(name_pattern, text)
    return emails, phones, tags

def upload_files(request):
    if request.method == 'POST' and request.FILES.getlist('pdf_files'):
        pdf_files = request.FILES.getlist('pdf_files')
        extracted_data = []
        with tempfile.TemporaryDirectory() as temp_dir:
            for pdf_file in pdf_files:
                file_path = os.path.join(temp_dir, pdf_file.name)
                with open(file_path, 'wb') as f:
                    for chunk in pdf_file.chunks():
                        f.write(chunk)
                try:
                    if pdf_file.name.lower().endswith("doc") or pdf_file.name.lower().endswith("docx") :
                        text = docx2txt.process(file_path)
                    elif pdf_file.name.lower().endswith("pdf") : 
                        text = extract_text_from_pdf(file_path)
                    else :
                        raise ValueError("unsupported file type ")
                    email , phone , tags = extract_info_from_text(text)
                    d = [email , phone , tags ]
                    extracted_data.append(d)
                except Exception as e:
                    print("error" , e)
                    continue 
                    # return render(request, 'error.html')
        headers = ["Email" , "Phone Number" , "Other Tags"] 
        output_file = generate_excel(headers , extracted_data)
        response = HttpResponse(output_file, content_type='application/vnd.ms-excel')
        response['Content-Disposition'] = 'attachment; filename="data.xls"'
        return response

        # print(extracted_data)
        # return render(request, 'results.html', {'extracted_data': extracted_data})

    return render(request, 'error.html')

def generate_excel(headers, data):

    wb = xlwt.Workbook()
    ws = wb.add_sheet('Data')

    for col, header in enumerate(headers):
        ws.write(0, col, header)

    for row, row_data in enumerate(data, start=1):
        for col, value in enumerate(row_data):
            ws.write(row, col, value)

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return output.getvalue()