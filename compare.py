import sys
from xml.dom import minidom

if len(sys.argv)!=3:
    print("usage: python compare2.py file1.xml file2.xml")
else:
    file1 = minidom.parse(sys.argv[1])
    print("loaded " + sys.argv[1])
    file2 = minidom.parse(sys.argv[2])
    print("loaded " + sys.argv[2])

    document1 = file1.getElementsByTagName('Document')
    print("got docsId in file1")
    document2 = file2.getElementsByTagName('Document')
    print("got docsId in file2")

    print("length of doc1: " + str(len(document1)))
    print("length of doc2: " + str(len(document2)))
    same = 0

    if len(document1) == len(document2):
        for elem1 in document1:
            #print(elem1.attributes['DocID'].value)
            found = False
            for elem2 in document2:
                if elem1.attributes['DocID'].value == elem2.attributes['DocID'].value:
                    #print(elem1.attributes['DocID'].value + " is equal " + elem2.attributes['DocID'].value +"\n")
                    same += 1
                    if same % 500 == 0:
                        print("Match amount: " + str(same))
                    found = True
                    break
            if found == False:
                print("they are not the same as cannot find a file1 docID in file2")
                break
    else:
        print("they are not the same as the number of the docID are different")

    if same == len(document1):
        print("these docs are the same")
    else:
        print("they are not the same")