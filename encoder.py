from lxml import etree
import xml.etree.ElementTree as ET
import array

class MisMatchexception(Exception):
   """Base class for other exceptions"""
   pass

PREFIX_WORD_PROC = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
PARAGRAPH_TAG = PREFIX_WORD_PROC + "p"
RUN_ELEMENT_TAG = PREFIX_WORD_PROC + "r"
RUN_ELEM_PROPERTY_TAG = PREFIX_WORD_PROC + "rPr"
BODY_TAG = PREFIX_WORD_PROC + "body"
TEXT_TAG = PREFIX_WORD_PROC + "t"
SZCS_TAG = PREFIX_WORD_PROC + "szCs"


tree = etree.parse('minimal/word/document.xml')
root = tree.getroot()


def merge_possible_run_elements(paragraph):
     #ricerca run elements nel paragrafo
     run_property_elements = paragraph.findall("./" + RUN_ELEMENT_TAG + "/" + RUN_ELEM_PROPERTY_TAG)
     run_elements_to_merge = []
     i = 0
     for node in run_property_elements:
        mismatch = False
        j = i + 1
        while mismatch != True and j < len(run_property_elements):
            for child_of_node in node:
                child_of_node_j = run_property_elements[j].find("./"  + child_of_node.tag)
                if child_of_node_j == None or (child_of_node.tag!= SZCS_TAG and child_of_node_j.attrib != child_of_node.attrib):
                    mismatch = True
                    break
            if(mismatch != True):
                j +=1
            else:
                j -=1
        x=i
        while x < j:
            run_elements_to_merge.append(paragraph.find("./" + RUN_ELEMENT_TAG + "[" + (x +1).__str__() + "]"))
            x +=1
        print(run_elements_to_merge)
        if(j== len(run_property_elements)):
            break
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





