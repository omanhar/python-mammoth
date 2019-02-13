# coding: utf8
import sys
import mammoth
from mammoth import documents


def style_to_str(style):
    ret = []
    for key in style:
        ret.append('%s: %s' % (key, style[key]))
    return ';'.join(ret)


def transform_items(element):
    style = {}
    if isinstance(element, documents.Paragraph):
        if element.alignment:
            style['text-align'] = element.alignment
    elif isinstance(element, documents.Run):
        if element.xml_properties:
            color = element.xml_properties.find_child('w:color')
            if color:
                style['color'] = '#%s' % color.attributes['w:val']
            size = element.xml_properties.find_child('w:szCs')
            if size:
                style['font-size'] = '%dpx' % round(int(size.attributes['w:val']) / 1.5)
    if style:
        element.html_attributes['style'] = style_to_str(style)
        return element.copy(html_attributes=element.html_attributes)
    return element


if __name__ == "__main__":
    fn = sys.argv[1]
    with open(fn, "rb") as docx:
        transform_document = mammoth.transforms.element_of_type((documents.Paragraph, documents.Run), transform_items)
        result = mammoth.convert_to_html(docx, transform_document=transform_document)

    f = open('%s.html' % fn, 'w')
    f.write(result.value.encode('utf8'))
    f.close()
