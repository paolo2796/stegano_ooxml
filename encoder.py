from lxml import etree
import copy
import random
import utils


PREFIX_WORD_PROC = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
PARAGRAPH_TAG = PREFIX_WORD_PROC + "p"
RUN_ELEMENT_TAG = PREFIX_WORD_PROC + "r"
RUN_ELEM_PROPERTY_TAG = PREFIX_WORD_PROC + "rPr"
BODY_TAG = PREFIX_WORD_PROC + "body"
TEXT_TAG = PREFIX_WORD_PROC + "t"
SZCS_TAG = PREFIX_WORD_PROC + "szCs"



ENCODED_INFORMATION = "empty"
ENCODED_INFORMATION_BITS = utils.text_to_binary(ENCODED_INFORMATION)
def merge_possible_run_elements(paragraph):
     #ricerca run elements nel paragrafo
     run_property_elements = paragraph.findall("./" + RUN_ELEMENT_TAG + "/" + RUN_ELEM_PROPERTY_TAG)
     i = 0

     for node in run_property_elements:
        mismatch = False
        j = i + 1
        #check sul numero degli elementi. Se non hanno lo stesso numero di elementi -> no merge
        if  j < len(run_property_elements) and len(run_property_elements[i]) !=  len(run_property_elements[j]):
            continue

        while mismatch != True and j < len(run_property_elements):
            for child_of_node in node:
                child_of_node_j = run_property_elements[j].find("./"  + child_of_node.tag)
                if child_of_node_j == None or (child_of_node.tag != SZCS_TAG and child_of_node_j.attrib != child_of_node.attrib):
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


def check_if_available_space(index,paragraph):
    count_zero = 0
    j = i
    while ENCODED_INFORMATION_BITS[j % len(ENCODED_INFORMATION_BITS)] == "0":
        j += 1
        count_zero += 1
    if count < count_zero + 1:
        node = paragraph.find("./" + RUN_ELEMENT_TAG + "[" + (offset_run_elem).__str__() + "]")
        if node.find("./" + RUN_ELEM_PROPERTY_TAG + "/" + SZCS_TAG) != None:
            node.find("./" + RUN_ELEM_PROPERTY_TAG + "/" + SZCS_TAG).set(PREFIX_WORD_PROC + "val", "-1")
        else:
            etree.SubElement(node, SZCS_TAG)
            node.find("./" + SZCS_TAG).set(PREFIX_WORD_PROC + "val", "-1")
        return False

    return True

#START


ENCODED_INFORMATION = input("Inserisci il testo segreto:")
ENCODED_INFORMATION_BITS = utils.text_to_binary(ENCODED_INFORMATION)
print(ENCODED_INFORMATION_BITS)

tree = etree.parse('minimal/word/document.xml')
root = tree.getroot()
paragraphs = root.findall("./" + BODY_TAG + "/" + PARAGRAPH_TAG)
i = 0
# step 10 -> Repeat Step 2 to Step 9 until all paragraph elements have been addressed.
for paragraph in paragraphs:
    #step 3 -> If two or more adjacent run elements in P have the same attributes, merge these run elements
    merge_possible_run_elements(paragraph)

    tree.write("output.xml")
    run_elements = paragraph.findall("./" + RUN_ELEMENT_TAG)
    i_run_elements = 1
    offset_run_elem = 1

    # rimuovo tutti i nodi != RUN_ELEMENT_TAG e li memorizzo in un array
    arr_childs_par_to_save = []
    for child_paragraph in paragraph.findall("./"):
        if child_paragraph.tag != RUN_ELEMENT_TAG:
            arr_childs_par_to_save.append(child_paragraph)
            paragraph.remove(child_paragraph)


    #step 4 -> Extract a run element <w:r>…</w:r> to R, corresponding text element <w:t>…</w:t> to T, and the corresponding attributes to A.
    while i_run_elements <= len(run_elements):
        N = 1
        #se non contiene il tag <w:t> -> aggiungilo al run element corrente
        if run_elements[i_run_elements - 1].find("./" + TEXT_TAG) == None:
            run_elements[i_run_elements - 1].append(etree.Element(TEXT_TAG))
            run_elements[i_run_elements - 1].find("./" + TEXT_TAG).text=""

        tag_element = run_elements[i_run_elements - 1].find("./" + TEXT_TAG)
        #se non contiene il tag szCS lo aggiungo al run element corrente
        if run_elements[i_run_elements - 1].find("./" + RUN_ELEM_PROPERTY_TAG + "/" + SZCS_TAG) == None:
            run_elements[i_run_elements - 1].find("./" + RUN_ELEM_PROPERTY_TAG).append(etree.Element(SZCS_TAG))
            run_elements[i_run_elements - 1].find("./" + RUN_ELEM_PROPERTY_TAG + "/" + SZCS_TAG).set(PREFIX_WORD_PROC + "val",random.randint(1, 10).__str__())

        #step 5 ->Make a count N=1 to accumulate the number of characters to be divided. Count the number of characters in T, and record it to C.
        count = len(tag_element.text)
        while count >=1:

            #check if enough space to inject information coded remain
            if(check_if_available_space(i,paragraph) == False):
                break

            #step 6 -> Read one bit from the encoded information H circularly, decrease the value of C by one (in this process, assume the digit “1” is used for splitting):
            # case a
            if(ENCODED_INFORMATION_BITS[i % len(ENCODED_INFORMATION_BITS)]) == "0":
                N += 1
            #case b
            else:
                tag_element = paragraph.find("./" + RUN_ELEMENT_TAG + "[" + (offset_run_elem).__str__() + "]" + "/" + TEXT_TAG)
                text = tag_element.text
                new_run_elem = copy.copy(paragraph.find("./" + RUN_ELEMENT_TAG + "[" + (offset_run_elem).__str__() + "]"))
                # step 8 -> Modify the splitting mark <w:szCs> in the run elements alternatively.
                if new_run_elem.find("./" + RUN_ELEM_PROPERTY_TAG + "/" + SZCS_TAG) != None:
                    new_run_elem.find("./" + RUN_ELEM_PROPERTY_TAG + "/" + SZCS_TAG).set(PREFIX_WORD_PROC + "val",random.randint(1,10).__str__())
                #aggiungo tag SZCS all'elemento rpr
                else:
                    new_run_elem.find("./" + RUN_ELEM_PROPERTY_TAG).insert(1,etree.Element(SZCS_TAG))
                    new_run_elem.find("./" + RUN_ELEM_PROPERTY_TAG + "/" + SZCS_TAG).set(PREFIX_WORD_PROC + "val",random.randint(1,10).__str__())

                tag_element.text = text[0:N]
                new_run_elem.find("./" + TEXT_TAG).text = text[N:]
                paragraph.insert(offset_run_elem,new_run_elem)
                offset_run_elem += 1
                tree.write("output.xml")
                N = 1
            i += 1
            count -= 1
            # step 7 -> Go to step 6 until the value of C is one.
        i_run_elements += 1
        offset_run_elem += 1
        # step 9 -> Repeat Step 4 to Step 8 until all run elements have been addressed in the paragraph element P
    # push di tutti i nodi != RUN_ELEMENT_TAG, memorizzati in arr_childs_par_to_save
    for item in arr_childs_par_to_save:
        paragraph.append(item)


tree.write("output.xml")
exit(0)







