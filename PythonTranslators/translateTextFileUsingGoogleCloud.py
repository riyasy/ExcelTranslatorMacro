import os
import re
import time

from google.cloud import translate_v2 as translate

os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = 'eighth-epsilon-313100-674cc277361c.json'

translate_client = translate.Client()


# text = 'Good Morning. GoodBye. And Hello'
# target = 'ja'
# output = translate_client.translate(
#     text,
#     target_language=target)
# print(output)


def check_if_jap(text: str):
    regex = u'[\u3040-\u30ff\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff\uff66-\uff9f]'
    # regex = u'[\p{Hiragana}\p{Katakana}\p{Han}]+'
    p = re.compile(regex, re.U)
    if p.search(text):
        return True
    else:
        return False


def translate():
    input_file = 'D:\\Translation\Input\Out.txt'
    output_file = 'D:\\Translation\Input\Out_Translated.txt'

    tobe_translated_list = []
    translated_file_content = []

    input_file = open(input_file, mode='r', encoding='utf-16')
    file_lines = input_file.readlines()
    for line in file_lines:
        line_text = line.strip()
        if (line_text != '~~~') & (check_if_jap(line_text) == True):
            tobe_translated_list.append(line_text)

    # output_file = open('D:\\Translation\Input\Out_Translated_BefProcessing.txt', mode='w', encoding='utf-16')
    # output_file.write('\n'.join(tobe_translated_list))
    # output_file.close()

    len_translated = 0
    list_of_list = []
    sublist = []
    for line in tobe_translated_list:
        # len_translated += sys.getsizeof(line)
        len_translated += len(line.encode('utf-8'))
        if len_translated > 4000:
            list_of_list.append(sublist[:])
            len_translated = 0
            sublist.clear()
        sublist.append(line)
    list_of_list.append(sublist[:])

    translated_list = []

    print_count = 0;
    for sublist1 in list_of_list:
        package = "\n".join(sublist1)
        if len(package.encode('utf-8')) > 4900:
            split_lines = package.split('\n')
        else:
            translated_package = translate_client.translate(package, target_language='en', source_language='ja',
                                                            format_='text')
            time.sleep(1)
            split_lines = translated_package['translatedText'].split('\n')
        for line in split_lines:
            print_count += 1
            translated_list.append(line)
            print("Line{}: {}".format(print_count, line))

    # output_file = open('D:\\Translation\Input\Out_Translated_AftProcessing.txt', mode='w', encoding='utf-16')
    # output_file.write('\n'.join(translated_list))
    # output_file.close()

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
    output_file.close()


translate()
