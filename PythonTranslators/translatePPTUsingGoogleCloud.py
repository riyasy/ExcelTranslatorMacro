import win32com.client
import os
from google.cloud import translate_v2 as translate

msoTypes = {}
msoTypes[30] = 'mso3DModel'
msoTypes[1]  = 'msoAutoShape'
msoTypes[2]  = 'msoCallout'
msoTypes[20] = 'msoCanvas'
msoTypes[3]  = 'msoChart'
msoTypes[4]  = 'msoComment'
msoTypes[27] = 'msoContentApp'
msoTypes[21] = 'msoDiagram'
msoTypes[7]  = 'msoEmbeddedOLEObject'
msoTypes[8]  = 'msoFormControl'
msoTypes[5]  = 'msoFreeform'
msoTypes[28] = 'msoGraphic'
msoTypes[6]  = 'msoGroup'
msoTypes[24] = 'msoIgxGraphic'
msoTypes[22] = 'msoInk'
msoTypes[23] = 'msoInkComment'
msoTypes[9]  = 'msoLine'
msoTypes[31] = 'msoLinked3DModel'
msoTypes[29] = 'msoLinkedGraphic'
msoTypes[10] = 'msoLinkedOLEObject'
msoTypes[11] = 'msoLinkedPicture'
msoTypes[16] = 'msoMedia'
msoTypes[12] = 'msoOLEControlObject'
msoTypes[13] = 'msoPicture'
msoTypes[14] = 'msoPlaceholder'
msoTypes[18] = 'msoScriptAnchor'
msoTypes[-2] = 'msoShapeTypeMixed'
msoTypes[19] = 'msoTable'
msoTypes[17] = 'msoTextBox'
msoTypes[15] = 'msoTextEffect'
msoTypes[26] = 'msoWebVideo'

tx_list = {}
os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = 'eighth-epsilon-313100-674cc277361c.json'

translate_client = translate.Client()





def open_presentation(filename):

    app = win32com.client.Dispatch("PowerPoint.Application")
    
    # app.Visible = True
    presentation = app.Presentations.Open(filename)
    return app, presentation


def save_presentation(ppt):
    ppt.Save()
    ppt.Close()
    


def process_slides(slide):
    process_shapes(slide.Shapes)
    # if slide.has_notes_slide:
        # print('----Notes----')
        # process_notes(slide.notes_slide)


def process_shapes(shapes):
    for shape in shapes:
        name = shape.name
        type = msoTypes[shape.type]
        print (type + ' : ' + name)
        
        if shape.TextFrame2.HasText:
            text = shape.TextFrame2.TextRange.text
            shape.TextFrame2.TextRange.text = translate(text)
        elif type == 'msoTable':
            process_table(shape.Table)
        elif type == 'msoChart':
            process_chart(shape.Chart)
        elif type == 'msoIgxGraphic':
            process_smartart(shape.SmartArt.Nodes)
        else:
            print ('Shape not processed:')
            print ('\t' + type + ' : ' + name)


def process_table(table):
    for row in range(1, table.rows.count + 1):
        for col in range(1, table.columns.count + 1):
            if table.cell(row, col).Shape.TextFrame2.HasText:
                text = table.cell(row, col).Shape.TextFrame2.TextRange.Text
                table.cell(row, col).Shape.TextFrame2.TextRange.Text = translate(text)


def process_chart(chart):   # TODO
    if chart.HasTitle:
        text = chart.ChartTitle.Text
        chart.ChartTitle.Text = translate(text)
    print (chart.SeriesCollection(1).XValues)
    xvalues = chart.SeriesCollection(1).XValues
    print (type(xvalues))
    tx_xvalues = []
    for value in xvalues:
        tx_xvalues.append(translate(value))
    chart.SeriesCollection(1).XValues = tuple(tx_xvalues)
    print (tuple(tx_xvalues))
    print ((chart.SeriesCollection))
    print (chart.SeriesCollection(1).Name)
    print (chart.SeriesCollection(1).XValues)
    print (chart.SeriesCollection(1).Values)
    print (chart.SeriesCollection(2).Name)
    print (chart.SeriesCollection(2).XValues)
    print (chart.SeriesCollection(2).Values)
    # if chart.HasLegend:
        # for legend in range(chart.SeriesCollection.count):
            # print (legend)


def process_smartart(nodes):
    for idx in range (1, nodes.count + 1):
        item = nodes(idx)
        if item.Nodes.Count > 0:
            process_smartart(item.Nodes)
        if item.TextFrame2.HasText:
            text = item.TextFrame2.TextRange.Text
            item.TextFrame2.TextRange.Text = translate(text)


def process_notes(notes):
    print (notes.notes_text_frame.text)
    # print (help(notes))


def translate(text):
    # if whole text is in english do not translate
    if not translation_required(text):
        return text

    # return local copy if already translated
    if text in tx_list.keys():
        return tx_list[text]

    # # FOR TESTING - TO REMOVE   # TODO
    # tx_text = 'TestingTx ' + text
    # return tx_text
    # # FOR TESTING - TO REMOVE   # TODO

    # target = 'en'
    tx_text = translate_client.translate(
        text,
        target_language='en', source_language='ja'
        )
    # print(tx_text)
    return tx_text['translatedText']


def translation_required(text):
    for ch in text:
        if ord(ch) > 255:
            return True
    return False


if __name__ == '__main__':
    scritpath = os.path.realpath(__file__)
    path = scritpath[:scritpath.rfind('\\')]
    files = os.listdir(path)
    filelist = []
    for fl in files:
        if fl.endswith('.pptx') or fl.endswith('.ppt') or fl.endswith('.pptm'):
            filelist.append(path + '/' + fl)
    for fl in filelist:
        app, ppt = open_presentation(fl)

        for slide in ppt.Slides:
            process_slides(slide)

        save_presentation(ppt)
