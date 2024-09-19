from flask import Flask, request, send_file, render_template, jsonify
from flask_cors import CORS
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import zipfile
from io import BytesIO

app = Flask(__name__)
CORS(app)


CERTIFICATE_TEMPLATE = 'certificate_template.png'
FONT_PATH = 'OpenSans-Italic-VariableFont_wdth,wght.ttf'

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    try:

        if 'file' not in request.files:
            return 'No file uploaded', 400

        file = request.files['file']
        if file.filename == '':
            return 'No selected file', 400


        try:
            data = pd.read_excel(file)
        except Exception as e:
            return 'Invalid Excel file format', 400

        data.columns = data.columns.str.strip()


        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, 'a', zipfile.ZIP_DEFLATED) as zf:
            for index, row in data.iterrows():
                student_name = row['Name']
                student_class = str(row['Class'])
                student_school = row['School']


                img = Image.open(CERTIFICATE_TEMPLATE)
                draw = ImageDraw.Draw(img)
                font = ImageFont.truetype(FONT_PATH, 45)


                name_position = (508, 306.5)
                text_position = (1032, 803)
                next_position = (841, 900)


                participation_text = f"{student_name} from {student_class} and {student_school}"
                next_line = "participated in the Career Guidance Program organized by Graphic Era Hill University, Bhimtal Campus. We wish you the best for your future!"


                draw.text(text_position, participation_text, font=font, fill="black")
                draw.text(next_position, next_line, font=font, fill="black")


                img_buffer = BytesIO()
                img.save(img_buffer, format='PNG')
                img_buffer.seek(0)


                zf.writestr(f'certificate_{student_name}.png', img_buffer.read())

        zip_buffer.seek(0)


        return send_file(zip_buffer, as_attachment=True, download_name='certificates.zip', mimetype='application/zip')

    except Exception as e:
        return jsonify({'error': 'An unexpected error occurred'}), 500

if __name__ == '__main__':
    app.run(debug=True)
