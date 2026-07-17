from fastapi import FastAPI, Request
import io
import os
import tempfile
import requests
import pdfplumber

from dotenv import load_dotenv
from requests.auth import HTTPBasicAuth

from docx import Document
from docx.document import Document as _Document
from docx.table import Table
from docx.text.paragraph import Paragraph


load_dotenv()

app = FastAPI()




baseURL = os.getenv(
    "FUSION_URL",
    "https://iaaley-test.fa.ocs.oraclecloud.com"
)

username = os.getenv("username")
password = os.getenv("password")




def iter_block_items(parent):

    if isinstance(parent, _Document):
        parent_elm = parent.element.body
    else:
        parent_elm = parent._tc


    for child in parent_elm.iterchildren():

        if child.tag.endswith("}p"):
            yield Paragraph(child, parent)

        elif child.tag.endswith("}tbl"):
            yield Table(child, parent)



def extract_docx(file_bytes):

    doc = Document(io.BytesIO(file_bytes))

    output = []


    for block in iter_block_items(doc):

        # Paragraph extraction

        if isinstance(block, Paragraph):

            text = block.text.strip()

            if text:
                output.append(text)


        # Table extraction

        elif isinstance(block, Table):

            output.append("")


            for row in block.rows:

                row_data = []

                for cell in row.cells:

                    value = (
                        cell.text
                        .replace("\n", " ")
                        .strip()
                    )

                    row_data.append(value)


                output.append(
                    " | ".join(row_data)
                )


            output.append("")


    return "\n".join(output)




def extract_pdf(file_bytes):

    output = []


    with pdfplumber.open(
        io.BytesIO(file_bytes)
    ) as pdf:


        for page in pdf.pages:


            text = page.extract_text()


            if text:
                output.append(text)



            tables = page.extract_tables()


            if tables:


                for table in tables:


                    output.append("")


                    for row in table:


                        if row:


                            cells = []


                            for cell in row:

                                if cell:
                                    cells.append(
                                        str(cell)
                                        .replace("\n"," ")
                                        .strip()
                                    )

                                else:
                                    cells.append("")


                            output.append(
                                " | ".join(cells)
                            )


                    output.append("")


    return "\n".join(output)




def extract_doc(file_bytes):

    with tempfile.TemporaryDirectory() as temp:


        doc_path = os.path.join(
            temp,
            "resume.doc"
        )


        with open(
            doc_path,
            "wb"
        ) as f:

            f.write(file_bytes)



        os.system(
            f"libreoffice --headless "
            f"--convert-to docx "
            f"{doc_path} "
            f"--outdir {temp}"
        )



        docx_path = os.path.join(
            temp,
            "resume.docx"
        )



        with open(
            docx_path,
            "rb"
        ) as f:

            return extract_docx(
                f.read()
            )




def get_resume_text(job_application_id):


    try:


        url =f"{baseURL}/hcmRestApi/resources/11.13.18.05/recruitingJobApplications/{job_application_id}/child/attachments"


        response = requests.get(

            url,

            auth=HTTPBasicAuth(
                username,
                password
            ),

            headers={
                "Accept":"application/json"
            }

        )



        response.raise_for_status()



        data = response.json()



        items = data.get(
            "items",
            []
        )



        if not items:

            return "No Attachment Found"



        # Find Resume file

        attachment = None



        for item in items:


            file_name = (

                item.get("FileName")
                or item.get("Title")
                or ""

            ).lower()



            if (
                "resume" in file_name
                or "cv" in file_name
            ):

                attachment = item

                break



        if attachment is None:

            attachment = items[0]



        file_name = (

            attachment.get("FileName")
            or attachment.get("Title")
            or ""

        )



        file_url = None



        for link in attachment.get(
            "links",
            []
        ):


            if link.get("name") == "FileContents":

                file_url = link.get(
                    "href"
                )

                break



        if not file_url:

            return "File URL Not Found"



        file_response = requests.get(

            file_url,

            auth=HTTPBasicAuth(
                username,
                password
            )

        )



        file_response.raise_for_status()



        file_bytes = file_response.content



        extension = os.path.splitext(
            file_name
        )[1].lower()



        # Detect file type

        if not extension:


            magic = file_bytes[:8]



            if magic.startswith(
                b"%PDF"
            ):

                extension = ".pdf"



            elif magic.startswith(
                b"PK"
            ):

                extension = ".docx"



            elif magic.startswith(
                b"\xd0\xcf\x11\xe0"
            ):

                extension = ".doc"



        print(
            "Detected:",
            extension
        )



        if extension == ".pdf":

            return extract_pdf(
                file_bytes
            )



        elif extension == ".docx":

            return extract_docx(
                file_bytes
            )



        elif extension == ".doc":

            return extract_doc(
                file_bytes
            )



        else:

            return "Unsupported File"



    except Exception as e:

        return f"Error: {str(e)}"



@app.post("/extract-resume-document")
async def extract_resume_document(
    request: Request
):


    body = await request.json()



    job_application_id = body.get(
        "Job_Application_Id"
    )



    if not job_application_id:

        return {

            "error":
            "Job_Application_Id is required"

        }



    resume_text = get_resume_text(
        job_application_id
    )



    return {


        "Job_Application_Id":
            job_application_id,


        "ResumeText":
            resume_text

    }