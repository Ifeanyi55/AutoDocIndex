# pip install stanza PyPDF2 python-docx

from collections import defaultdict
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches
import stanza
import PyPDF2
import re

# stanza.download('en',model_dir='/content/sample_data')

# initialize the pipeline
nlp = stanza.Pipeline('en', processors='tokenize,ner')

def autodocIndex(doc,save_as):
  # read pdf document
  pdf_doc = PyPDF2.PdfReader(doc)

  pdf_text = []
  # extract all the pages of the document and store in a list
  for page in pdf_doc.pages:
    pdf_text.append(page.extract_text())

  # clean the text by removing line break symbol
  pdf_text = [re.sub(r'\s+', ' ', text) for text in pdf_text]

  # extract all the entities in the text and svae in respective lists
  doc_entities = [[ent.text for ent in nlp(pdf_text[i]).ents if ent.type in ['PERSON', 'ORG', 'GPE', 'LOCATION']] for i in range(len(pdf_text))]

  # append page numbers to each index
  doc_entities_labelled = [[f"{ele}     {num + 1}"  for ele in keywords] for num, keywords in enumerate(doc_entities)]

  # merge sublists into one index list
  index = [entities for i in doc_entities_labelled for entities in i]

  # arrange index in alphabetical order
  index.sort()

  # remove repeated elements and group page numbers together
  grouped = defaultdict(list)
  seen = defaultdict(set)

  for item in index:
    name, number = item.strip().rsplit(" ", 1)
    if number not in seen[name]:
        grouped[name].append(number)
        seen[name].add(number)

  # Format the result
  result = [f"{name} {','.join(numbers)}" for name, numbers in grouped.items()]

  # create a Word document
  document = Document()

  # set two-column layout
  for section in document.sections:
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)
    section.top_margin = Inches(0.75)
    section.bottom_margin = Inches(0.75)

    # get section properties and configure 2 columns
    sectPr = section._sectPr
    cols = sectPr.first_child_found_in("w:cols")
    if cols is None:
        cols = OxmlElement(qn("w:cols"))
        sectPr.append(cols)
    cols.set(qn("w:num"), "2")
    cols.set(qn("w:space"), "720")  # optional spacing between columns (in twips)

    # write content, insert page break every 40 lines
    for i, item in enumerate(result):
        para = document.add_paragraph(item)
        para.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

        if (i + 1) % 40 == 0:
            document.add_page_break()

    # save document
    document.save(save_as + ".docx")

if __name__ == "__main__":
  autodocIndex(doc="memoir.pdf",save_as="memoir_index")
  print("Memoir Index Created")
