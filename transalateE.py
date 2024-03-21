from zipfile import ZipFile, ZIP_DEFLATED
from lxml import etree
from lxml.etree import Element
from transformers import AutoTokenizer, AutoModelForSeq2SeqLM
import io
model_name = "VietAI/envit5-translation"
tokenizer = AutoTokenizer.from_pretrained(model_name)
model = AutoModelForSeq2SeqLM.from_pretrained(model_name)

CONTENT_TYPE = "http://schemas.openxmlformats.org/package/2006/content-types"

CONTENT_TYPES_PARTS = (
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',  # noqa
    'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',  # noqa
    'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml',  # noqa
)

path_file = 'D:\\test\\b\\file_run\\translate\\a.docx'

file_zip = ZipFile(path_file)

# print(file_zip.open('[Content_Types].xml'))3
# truy suất đến phần tử w:...
# w: http://schemas.openxmlformats.org/wordprocessingml/2006/main
# './/{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tenphantu'

part__path_file = {}
path_name = ""
header_part = ""
body_part = ""
footer_part = ""
# new_merge_field.set('is_merge_field', 'True'): set new atrribute
NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


for file in etree.parse(file_zip.open('[Content_Types].xml')).findall(f'{{{CONTENT_TYPE}}}Override'):
    full_path_name = file.attrib["PartName"]
    content_type = file.attrib['ContentType']
    if content_type in CONTENT_TYPES_PARTS:
        path_name = full_path_name[1:]
        print(path_name)
        part__path_file[path_name] = etree.parse(file_zip.open(path_name))
one_string = ""
inputs = [

]

# chỉ khi nào cả câu đều cùng 1 dạng thì mới lấy nguyên câu và định dạng đó cho vào word sau khi translate
index = 0
parent_w_p = parent_w_p_previos = None
for path in part__path_file[path_name].findall(f'.//{{{NAMESPACE}}}t'):
    data = path.text
    print(data)

    if parent_w_p:
        print()
        parent_w_p_previos = parent_w_p

    parent = path.getparent()
    # parent_w_p = parent.getparent()
    # print("aaaaaaaaaaaaaaaa", etree.tostring(parent.getnext()))
    if  parent.getnext() is None:
        print("da het dong")
    if parent_w_p_previos:
        print("vooooooooooo")
        print(parent_w_p_previos != parent_w_p)
        # print(etree.tostring(parent_w_p_previos),  etree.tostring(parent_w_p))
    one_string += data if not data.isspace() else ""
    if "." in path.text or "…" in path.text or parent.getnext() is None:
        get_previous_tag = parent.getprevious()
        have_style = get_previous_tag.find(f'.//{{{NAMESPACE}}}rPr')
        if parent.find(f'.//{{{NAMESPACE}}}rPr') is not None:
            path.set("index", str(index))
        elif have_style is not None:
            print("voô eoiii")
            path_w_t_in_previous = get_previous_tag.find(f'.//{{{NAMESPACE}}}t')
            path_w_t_in_previous.set("index", str(index))
            path.text = ""
        else:
            path.set("index", str(index))



        inputs.append(f"vi: {one_string}")
        index +=1
        # print("in", index)
        one_string = ""


    else:
        path.text = ""



print(inputs)
outputs = model.generate(tokenizer(inputs, return_tensors="pt", padding=True).input_ids.to('cpu'), max_length=512)
data_outputs = tokenizer.batch_decode(outputs, skip_special_tokens=True)
print(data_outputs)
for i in range(0, index):
    print(i)
    merge_fields = part__path_file[path_name].find(f'.//{{{NAMESPACE}}}t[@index="{i}"]')
    print("len", len(merge_fields))
    merge_fields.text = str(data_outputs[i][4:])
docx_bytes_io = io.BytesIO()

with ZipFile(docx_bytes_io, 'w') as output:
    for zipinfo in file_zip.filelist:
        filename = zipinfo.filename
        if filename in part__path_file:
            xml = etree.tostring(part__path_file[filename].getroot())
            output.writestr(filename, xml)
        else:
            output.writestr(filename, file_zip.read(zipinfo))

with open(r"D:\test\b\file_run\translate\conver_a.docx",  "wb") as f:
    f.write(docx_bytes_io.getvalue())

