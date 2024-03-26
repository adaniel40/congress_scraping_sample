import urllib.request
import pandas as pd
from pandas import ExcelWriter
import xlsxwriter
import re
import os


#Set active directory
os.chdir("/Users/[user]/Documents/")


#Set url of committee doc
fp = urllib.request.urlopen("https://www.govinfo.gov/content/pkg/CHRG-118hhrg53204/html/CHRG-118hhrg53204.htm")
mybytes = fp.read()

mystr = mybytes.decode("utf8")
fp.close()


def witness(wit_list):
    wit_temp = wit_list.partition("...")[0]
    wit_temp = wit_temp.replace("\n","")
    wit_temp = wit_temp.replace("   "," ")
    wit_temp = wit_temp.replace(" D.C"," D.C.")
    return(wit_temp)

def wit_break(wit_list,cutchar):
    temp = wit_list.partition(cutchar)[2]
    if temp.partition("\n")[2] == "":
        return(temp)
    else:
        while temp[0] == " " or temp[0] == ".":
            temp = temp.partition("\n")[2]
        return(temp)

def pos_fix(pos):
    counter = 0
    for x in pos:
        if x ==",":
                counter += 1
    return(counter)

def pos_group(pos_count, position):
    if pos_count == 2:
        temp1 = position.partition(",")[0]
        temp2 = position.partition(",")[2]
        temp3 = temp2.partition(",")[2]
        temp2 = temp2.partition(",")[0]

        r = temp1 + "," + temp2
        o = temp3.strip()
        return(r,o)
    else:
        r = position.partition(",")[0]
        o = position.partition(", ")[2]
        return(r,o)



def wit_clean(wit):
    #location
    n=1
    counter=0
    while counter<2:
        if wit.strip() == "":
            return("")
        elif wit[len(wit) - n] == ",":
            if counter == 1:
                loc_temp = wit[len(wit) - (n):]
                loc = loc_temp[2:]
                wit = wit.partition(loc_temp)[0]
                counter += 1
            else:
                counter += 1
                n += 1
        else:
            n += 1

    #last name
    lname = wit.partition(",")[0]
    wit = wit.partition(",")[2]

    #first name
    fname = wit.partition(",")[0]
    fname = fname.strip()
    wit = wit.partition(",")[2]

    #nickname
    if fname[len(fname)-1] == "'":
        nname = fname.partition(" ``")[2]
        nname = nname.partition("''")[0]
        fname = fname.partition(" ``")[0]
    else:
        nname = ""


    #position 1
    p1 = wit.partition(";")[0]
    p1 = p1.strip()

    pos_count = pos_fix(p1)

    r1 = pos_group(pos_count,p1)[0]
    o1 = pos_group(pos_count,p1)[1]

    #position 2
    wit = wit.partition("; ")[2]
    if wit == "":
        r2 = ""
        o2 = ""

    else:
        p2 = wit.partition(";")[0]
        p2 = p2.strip()
        pos_count = pos_fix(p2)
        r2 = pos_group(pos_count,p2)[0]
        o2 = pos_group(pos_count,p2)[1]


    return(loc,lname,fname,nname,r1, o1, r2, o2)

def witness_scrape(witblock):
    witness_list = []
    while witblock != "":
        wit_out = []
        wit_temp = witness(witblock)
        cutchar = wit_temp[len(wit_temp) - 5:]
        witblock = wit_break(witblock,cutchar)
        wit_out = wit_clean(wit_temp)
        witness_list.append(wit_out)

    return(witness_list)

def metadata(mystr):
    meta = []
    #pull out title
    temp = mystr.partition("<title> - ")[2]
    title = temp.partition("</title")[0]
    title = title.replace("\n", '')
    meta.append(title)

    # pull out committee
    temp = mystr.partition("COMMITTEE ON ")[2]
    tempcmte = temp.partition("\n")[0]
    tempcmte = tempcmte.lower()
    cmte = "committee on "
    cmte += tempcmte
    meta.append(cmte)

    # pull out chamber
    temp = mystr.partition("]")[0]
    cmbr = temp.partition("[")[2]
    cmbr = cmbr.partition(",")[0]
    meta.append(cmbr)

    # pull out congress
    temp = mystr.partition("]")[0]
    cong = temp.partition("Hearing, ")[2]
    meta.append(cong)

    # pull out date
    temp = mystr.partition("__________")[2]
    date = temp.partition("__________")[0]
    date = date.strip()
    meta.append(date)

    # pull out jacket num
    temp = mystr.partition(" WASHINGTON :")[0]
    jnum = temp.partition("PUBLISHING OFFICE")[2]
    jnum = jnum.strip()
    jnum = jnum.replace(" PDF", '')
    meta.append(jnum)

    return(meta)

def total(mystr):
    #pull out metadata
    meta = metadata(mystr)
    title = meta[0]

    # pull out witnesses
    temp = mystr.partition("Witnesses")[2]
    title_cut = title[0:10]
    witblock = temp.partition(title_cut)[0]
    witnesses = witness_scrape(witblock)

    return(meta,witnesses)






output = total(mystr)
file_name = output[0][5]+".xlsx"
meta = output[0]
df1 = pd.DataFrame([meta], columns = ["Title", "Committee", "Chamber", "Congress", "Date", "Jacket Number"])
print(df1)

wit = output[1]
df2 = pd.DataFrame(wit, columns= ["Location", "Last Name", "First Name", "Nickname", "Role 1", "Organization 1","Role 2", "Organization 2"])
print(df2)

with pd.ExcelWriter(file_name, engine='xlsxwriter') as writer:
    df2.to_excel(writer, sheet_name='Witnesses')
    df1.to_excel(writer, sheet_name='Metadata')

print("finished")
