from lxml import etree
import array


PREFIX_WORD_PROC = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
PARAGRAPH_TAG = PREFIX_WORD_PROC + "p"
RUN_ELEMENT_TAG = PREFIX_WORD_PROC + "r"
RUN_ELEM_PROPERTY_TAG = PREFIX_WORD_PROC + "rPr"
BODY_TAG = PREFIX_WORD_PROC + "body"


tree = etree.parse('minimal/word/document.xml')
root = tree.getroot()



def merge_possible_run_elements(paragraph):
     #ricerca run elements nel paragrafo
     run_property_elements = paragraph.findall("./" + RUN_ELEMENT_TAG + "/" + RUN_ELEM_PROPERTY_TAG)
     i = 0
     for node in run_property_elements:
         for child_of_node in node:
            mismatch = False
            if i + 1 < len(run_property_elements):
                child_of_node_next = run_property_elements[i + 1].find("./"  + child_of_node.tag)
                if child_of_node_next != None :
                    if child_of_node_next.attrib != child_of_node.attrib:
                        mismatch = True
                        break

         #endfor
         i +=1






paragraphs = root.findall("./" + BODY_TAG + "/" + PARAGRAPH_TAG)
for node in paragraphs:
    merge_possible_run_elements(node)
exit(0)

'''
for child in root:
    if child.tag == BODY_TAG:
        for child_of_body in child :
            if child_of_body.tag == PARAGRAPH_TAG :
                arr_run_elems_property = []
                for child_of_paragraph in child_of_body :
                    if(child_of_paragraph.tag == RUN_ELEMENT_TAG) :
                        for child_of_run_element in child_of_paragraph :
                            if(child_of_run_element.tag == RUN_ELEM_PROPERTY_TAG ) :
                                arr_run_elems_property.append(child_of_paragraph)
                # end  child_of_paragraph
                merge_possible_run_elements(arr_run_elems_property) 
'''





