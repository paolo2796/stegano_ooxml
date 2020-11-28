from lxml import etree
import copy
import random
import utils
from distutils.dir_util import copy_tree
from distutils.dir_util import copy_tree
import shutil
import gzip
import os
import zipfile





PREFIX_WORD_PROC = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
PARAGRAPH_TAG = PREFIX_WORD_PROC + "p"
RUN_ELEMENT_TAG = PREFIX_WORD_PROC + "r"
RUN_ELEM_PROPERTY_TAG = PREFIX_WORD_PROC + "rPr"
BODY_TAG = PREFIX_WORD_PROC + "body"
TEXT_TAG = PREFIX_WORD_PROC + "t"
SZCS_TAG = PREFIX_WORD_PROC + "szCs"

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
                    if node == None:
                        break
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
                    if node == None:
                        break
                    #NEW DA TESTARE COMPLETO
                    if paragraph.find("./" + RUN_ELEMENT_TAG + "[" + (i + 1).__str__() + "]" + "/" + TEXT_TAG) == None:
                        paragraph.find("./" + RUN_ELEMENT_TAG + "[" + (i + 1).__str__() + "]").append(etree.Element(TEXT_TAG))
                        paragraph.find("./" + RUN_ELEMENT_TAG + "[" + (i + 1).__str__() + "]" + "/" + TEXT_TAG).text = node.text
                    else:
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



def check_if_available_space(index,paragraph, information_to_encode_bits, offset_run_elem,i,count):
    count_zero = 0
    j = i
    while information_to_encode_bits[j % len(information_to_encode_bits)] == "0":
        j += 1
        count_zero += 1
    if count < count_zero + 1:
        node = paragraph.find("./" + RUN_ELEMENT_TAG + "[" + (offset_run_elem).__str__() + "]")
        if node.find("./" + RUN_ELEM_PROPERTY_TAG + "/" + SZCS_TAG) != None:
            node.find("./" + RUN_ELEM_PROPERTY_TAG + "/" + SZCS_TAG).set(PREFIX_WORD_PROC + "val", "#")
        else:
            etree.SubElement(node, SZCS_TAG)
            node.find("./" + SZCS_TAG).set(PREFIX_WORD_PROC + "val", "#")
        return False
    return True


def createFileStego(tree,name_file):
    tree.write("app/stego/document.xml")
    copy_tree("input/" + name_file + "/file_extracted" , "app/stego/file_extracted")
    shutil.copy("app/stego/document.xml", "app/stego/file_extracted/word")
    zf = zipfile.ZipFile("app/stego/stego.zip", "w",zipfile.ZIP_DEFLATED)
    for dirname, subdirs, files in os.walk("./app/stego/file_extracted"):
        for filename in files:
            path=""
            if dirname == "./app/stego/file_extracted":
                path = filename
            else:
                path = dirname.split("./app/stego/file_extracted/")[1] + "/" + filename
            zf.write(os.path.join(dirname, filename),path)
    zf.close()
    os.rename('./app/stego/stego.zip','./app/stego/stego.docx')
    os.remove('./app/stego/document.xml')
    return "app/stego/stego.docx"





def encoding(path,message,password):
    # step 1 -> Leggi il codice dal file "document.xml", relativo al documento D
    tree = etree.parse("input/" + path + '/file_extracted/word/document.xml')
    root = tree.getroot()
    total_counter_words = 0
    # step 2 -> Cifra il testo segreto H mediante l’algoritmo AES-CBC, usando la chiave simmetrica
    ENCRYPTED = utils.encrypt(password, message)
    # step 3 -> aggiungi alla fine del testo cifrato il carattere di divisione "%"
    INFORMATION_TO_ENCODE_BITS = utils.text_to_binary(ENCRYPTED.decode('utf-8')) + utils.text_to_binary(
        utils.MAGIC_CHAR_SPLIT)

    print("INIEZIONE IN CORSO .....")
    paragraphs = root.findall("./" + BODY_TAG + "/" + PARAGRAPH_TAG)
    i = 0
    # step 14 -> Ripeti step 6 - step 13 finché tutti gli elementi <w:r> in P  non sono stati risolti;
    for paragraph in paragraphs:
        # step 5 -> IF (2 o più elementi <w:r> consecutivi in P hanno gli stessi attributi) THEN unisci gli elementi consecutivi <w:r>.;
        merge_possible_run_elements(paragraph)

        run_elements = paragraph.findall("./" + RUN_ELEMENT_TAG)
        i_run_elements = 1
        offset_run_elem = 1

        # aggiungo tutti i nodi != RUN_ELEMENT_TAG e li memorizzo in un array
        arr_childs_par_to_save = []
        for child_paragraph in paragraph.findall("./"):
            if child_paragraph.tag != RUN_ELEMENT_TAG:
                arr_childs_par_to_save.append(child_paragraph)
                # paragraph.remove(child_paragraph)

        # step 14 -> risolvi tutti gli elementi <w:r> in P
        while i_run_elements <= len(run_elements):
            # step8 -> Inizializza il contatore N=1 per accumulare il numero di caratteri da dividere. Successivamente, conta il numero di caratteri in T, e memorizzali nella variabile C.
            N = 1
            # se non contiene il tag <w:t> -> aggiungilo al run element corrente
            if run_elements[i_run_elements - 1].find("./" + TEXT_TAG) == None:
                run_elements[i_run_elements - 1].append(etree.Element(TEXT_TAG))
                run_elements[i_run_elements - 1].find("./" + TEXT_TAG).text = ""
            total_counter_words += len(run_elements[i_run_elements - 1].find("./" + TEXT_TAG).text)

            tag_element = run_elements[i_run_elements - 1].find("./" + TEXT_TAG)
            # se non contiene il tag rpr lo aggiungo al run element corrente --> DA TESTARE COMPLETO
            if run_elements[i_run_elements - 1].find("./" + RUN_ELEM_PROPERTY_TAG) == None:
                rpr = etree.Element(RUN_ELEM_PROPERTY_TAG)
                rpr.append(etree.Element(SZCS_TAG))
                rpr.find("./" + SZCS_TAG).set(PREFIX_WORD_PROC + "val", random.randint(1, 300).__str__())
                run_elements[i_run_elements - 1].append(rpr)
            # se non contiene il tag szCS lo aggiungo al run element corrente
            elif run_elements[i_run_elements - 1].find("./" + RUN_ELEM_PROPERTY_TAG + "/" + SZCS_TAG) == None:
                run_elements[i_run_elements - 1].find("./" + RUN_ELEM_PROPERTY_TAG).append(etree.Element(SZCS_TAG))
                run_elements[i_run_elements - 1].find("./" + RUN_ELEM_PROPERTY_TAG + "/" + SZCS_TAG).set(
                    PREFIX_WORD_PROC + "val", random.randint(1, 300).__str__())

            # step 8 -> Inizializza il contatore N=1 per accumulare il numero di caratteri da dividere. Successivamente, conta il numero di caratteri in T, e memorizzali nella variabile C.
            count = len(tag_element.text)
            # step 12 -> Ritorna allo step 9 fino a quando C >= 1;
            while count >= 1:
                # check if enough space to inject information coded remain
                if (check_if_available_space(i, paragraph, INFORMATION_TO_ENCODE_BITS, offset_run_elem,i,count) == False):
                    break
                # step 9,10,11
                # case a
                if (INFORMATION_TO_ENCODE_BITS[i % len(INFORMATION_TO_ENCODE_BITS)]) == "0":
                    N += 1
                # case b
                else:
                    tag_element = paragraph.find(
                        "./" + RUN_ELEMENT_TAG + "[" + (offset_run_elem).__str__() + "]" + "/" + TEXT_TAG)
                    text = tag_element.text
                    new_run_elem = copy.copy(
                        paragraph.find("./" + RUN_ELEMENT_TAG + "[" + (offset_run_elem).__str__() + "]"))
                    # step 8 -> Modify the splitting mark <w:szCs> in the run elements alternatively.
                    if new_run_elem.find("./" + RUN_ELEM_PROPERTY_TAG + "/" + SZCS_TAG) != None:
                        new_run_elem.find("./" + RUN_ELEM_PROPERTY_TAG + "/" + SZCS_TAG).set(PREFIX_WORD_PROC + "val",
                                                                                             random.randint(1,
                                                                                                            300).__str__())
                    # aggiungo tag SZCS all'elemento rpr
                    else:
                        new_run_elem.find("./" + RUN_ELEM_PROPERTY_TAG).insert(1, etree.Element(SZCS_TAG))
                        new_run_elem.find("./" + RUN_ELEM_PROPERTY_TAG + "/" + SZCS_TAG).set(PREFIX_WORD_PROC + "val",
                                                                                             random.randint(1,
                                                                                                            300).__str__())

                    tag_element.text = text[0:N]
                    new_run_elem.find("./" + TEXT_TAG).text = text[N:]
                    # aggiungo space preserve
                    tag_element.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                    new_run_elem.find("./" + TEXT_TAG).set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                    paragraph.insert(paragraph.find("./" + RUN_ELEMENT_TAG + "[" + (offset_run_elem).__str__() + "]").getparent().index(
                        paragraph.find("./" + RUN_ELEMENT_TAG + "[" + (offset_run_elem).__str__() + "]")) + 1,
                                     new_run_elem)
                    offset_run_elem += 1
                    tree.write("stego/document.xml")
                    N = 1
                i += 1
                count -= 1
                # step 7 -> Go to step 6 until the value of C is one.
            i_run_elements += 1
            offset_run_elem += 1
            # step test_per_la_scalabilita -> Repeat Step 4 to Step 8 until all run elements have been addressed in the paragraph element P
        # push di tutti i nodi != RUN_ELEMENT_TAG, memorizzati in arr_childs_par_to_save
    #  for item in arr_childs_par_to_save:
    #      paragraph.append(item)
    total_counter_inclusion = i

    if len(INFORMATION_TO_ENCODE_BITS) > total_counter_inclusion:
        print("non è stato possibile iniettare il testo segreto poichè presenta un numero di bits maggiori della capacità di inclusione")
    path = createFileStego(tree,path)
    return path









