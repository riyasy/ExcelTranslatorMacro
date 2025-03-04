import re
from googletrans import Translator

def check_if_jap(text: str):
    regex = u'[\u3040-\u30ff\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff\uff66-\uff9f]'
    p = re.compile(regex, re.U)
    if p.search(text):
        return True
    else:
        return False

def translate():
    input_file = 'D:\\Translation\Input\Out.txt'
    output_file = 'D:\\Translation\Input\Out_Translated.txt'

    translator = Translator()
    tobe_translated_list = []
    translated_file_content = []

    input_file = open(input_file, mode='r', encoding='utf-16')
    file_lines = input_file.readlines()
    for line in file_lines:
        line_text = line.strip()
        if (line_text != '~~~') & (check_if_jap(line_text) == True):
            tobe_translated_list.append(line_text)

    len_translated = 0
    list_of_list = []
    sublist = []
    for line in tobe_translated_list:
        len_translated += len(line)
        if len_translated > 4000:
            list_of_list.append(sublist[:])
            len_translated = 0
            sublist.clear()
        sublist.append(line)

    translated_list = []

    for sublist1 in list_of_list:
        translated_sublist = translator.translate(sublist1, src='ja', dest='en')
        for subtext in translated_sublist:
            translated_list.append(subtext.text)


    count = 0
    for line in file_lines:
        line_text = line.strip()
        if (line_text == '~~~') | (check_if_jap(line_text) == False):
            translated_file_content.append(line_text)
        else:
            translated_file_content.append(translated_list[count])
            count += 1

    output_file = open(output_file, mode='w', encoding='utf-16')
    output_file.write('\n'.join(translated_file_content))
    # output_file.writelines(translated_file_content)
    output_file.close()

translate()
