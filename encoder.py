from lxml import etree
import array


PREFIX_WORD_PROC = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
PARAGRAPH_TAG = PREFIX_WORD_PROC + "p"
RUN_ELEMENT_TAG = PREFIX_WORD_PROC + "r"
RUN_ELEM_PROPERTY_TAG = PREFIX_WORD_PROC + "rPr"
BODY_TAG = PREFIX_WORD_PROC + "body"


tree = etree.parse('minimal/word/document.xml')
root = tree.getroot()


def merge_possible_run_elements(items):
     mismatch = False
     for i, item in enumerate(items):
         for child_of_run_element in item:
            mismatch = False
            if child_of_run_element.tag == RUN_ELEM_PROPERTY_TAG :
                for child_of_rpr in child_of_run_element:
                    print(child_of_rpr)
            #endif
         '''
         
            for  child_of_run_element in item:
             mismatch = False
                 for j, child_of_rpr in child_of_run_element :
                     if child_of_rpr == items[i + 1][0][j] :
                         continue
                     else:
                         mismatch = True
                         break
             #endif     
         '''





paragraphs = root.findall("./" + BODY_TAG + "/" + PARAGRAPH_TAG + "[2]")
print(paragraphs)
exit(0)


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




