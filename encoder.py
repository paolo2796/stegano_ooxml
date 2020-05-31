from lxml import etree
import xml.etree.ElementTree as ET

import array
import copy


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


ENCODED_INFORMATION = ""
ENCODED_INFORMATION_BITS = (''.join(format(ord(x), '#010b') for x in ENCODED_INFORMATION))[2:]

tree = etree.parse('minimal/word/document.xml')
root = tree.getroot()


def merge_possible_run_elements(paragraph):
     #ricerca run elements nel paragrafo
     run_property_elements = paragraph.findall("./" + RUN_ELEMENT_TAG + "/" + RUN_ELEM_PROPERTY_TAG)
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
            #merge nodi fino al j - 1 elemento
            if mismatch == True:
                x = i + 1
                while x < j:
                    #append  node i + 1 to base node
                    node = paragraph.find("./" + RUN_ELEMENT_TAG + "[" + (x + 1).__str__() + "]" + "/" + TEXT_TAG)
                    base_node =  paragraph.find("./" + RUN_ELEMENT_TAG + "[" + (i + 1).__str__() + "]" + "/" + TEXT_TAG)
                    base_node.text = base_node.text + node.text
                    x += 1
                x = i + 1
                while x < j:
                    #rimuovo elemento successivo al nodo merge
                    run_property_elements.remove(run_property_elements[i + 1])
                    paragraph.remove(paragraph.find("./" + RUN_ELEMENT_TAG + "[" + (i + 2).__str__() + "]"))
                    x += 1
            #se sono arrivato alla fine dei nodi del paragrafo --> tutti i nodi hanno gli stessi attributi
            elif (j == len(run_property_elements) - 1):
                x = i + 1
                while x <= j:
                    #append  node i + 1 to base node
                    node = paragraph.find("./" + RUN_ELEMENT_TAG + "[" + (x + 1).__str__() + "]" + "/" + TEXT_TAG)
                    base_node =  paragraph.find("./" + RUN_ELEMENT_TAG + "[" + (i + 1).__str__() + "]" + "/" + TEXT_TAG)
                    base_node.text = base_node.text + node.text
                    x += 1
                x = i + 1
                while x <= j:
                    #rimuovo elemento successivo al nodo merge
                    run_property_elements.remove(run_property_elements[i + 1])
                    paragraph.remove(paragraph.find("./" + RUN_ELEMENT_TAG + "[" + (i + 2).__str__() + "]"))
                    x += 1
            else:
                j += 1

        if(j== len(run_property_elements)):
            break
        i +=1


def shift_run_element_by_pos(paragraph,start):
    count = len(paragraph.findall("./" + RUN_ELEMENT_TAG))
    while(count >= start):
        paragraph.insert(count + 1,copy.copy(paragraph.find("./" + RUN_ELEMENT_TAG + "[" + (count).__str__() + "]")))
        count -= 1


paragraphs = root.findall("./" + BODY_TAG + "/" + PARAGRAPH_TAG)
i = 0
for paragraph in paragraphs:
    merge_possible_run_elements(paragraph)
    run_elements = paragraph.findall("./" + RUN_ELEMENT_TAG)
    i_run_elements = 1
    offset_run_elem = 1
    while i_run_elements <= len(run_elements):
        N = 1
        tag_element = run_elements[i_run_elements - 1].find("./" + TEXT_TAG)
        count = len(tag_element.text)
        while count >=1:
            if(ENCODED_INFORMATION_BITS[i % len(ENCODED_INFORMATION_BITS)]) == "0":
                N += 1
            else:
                tag_element = paragraph.find("./" + RUN_ELEMENT_TAG + "[" + (offset_run_elem).__str__() + "]" + "/" + TEXT_TAG)
                text = tag_element.text
                new_run_elem = copy.copy(paragraph.find("./" + RUN_ELEMENT_TAG + "[" + (offset_run_elem).__str__() + "]"))
                paragraph.insert(offset_run_elem,new_run_elem)
                offset_run_elem += 1
                tag_element.text = text[0:N]
                new_run_elem.find("./" + TEXT_TAG).text = text[N:]
                N = 1
            i += 1
            count -= 1
        i_run_elements += 1
        offset_run_elem += 1
tree.write("output.xml")










exit(0)







