import urllib2
import re
import urllib




def find_between( s, first, last ):
    try:
        start = s.index( first ) + len( first )
        end = s.index( last, start )
        return s[start:end]
    except ValueError:
        return ""

def replace_all(text, dic):
    for i in dic:
        text = text.replace(i, '')
    return text

def replace_all_2(text, dic):
    for i,j in dic.iteritems():
        text = text.replace(i, j)
    return text


print "#################################################################################################"
print "#################################### CUSAT RESULT DOWNLOADER ####################################"
print "####################################       BY KISHOR V       ####################################"
print "#################################################################################################"
print "\n\nOpen a result on the webpage and input values as per the website"



#Variable initialisation
sem = raw_input("Semester : ")
course = "B.Tech" if (int(input("Degree [ 1. B.Tech | 2. M.Tech ] (Integer) : ")) == 1) else "M.Tech" 
exam = raw_input("Result type [Regular | Revaluation | Supplementary | etc.] : ")
month = raw_input("Month : ")
year = raw_input("Year : ")
i = int(input("Starting Reg. Number : "))
last = int(input("Last Reg. Number : "))

result = ""
branch = ""
url = "http://exam.cusat.ac.in/erp5/cusat/CUSAT-RESULT/Result_Declaration/display_sup_result?statuscheck=failed&deg_name="+course+"&semester="+urllib.quote(sem)+"&month="+month+"&year="+year+"&result_type="+exam+"&regno="
            
print "Downloading...\n\n"
while(i<=last):
    try:
        u = urllib2.urlopen(url+str(i))
        #Complete HTML
        text = u.read()
        #Replaced Line Breaks
        text = text.replace('\n','')
        #Searched for first table ( Name and details )
        data = find_between(text,'<table width="100%" class="order-list" border="3">','</table>')
        #Removed first table 
        text = text.replace('<table width="100%" class="order-list" border="3">','',1)
        #Searched for second table ( Marks )
        data = data + find_between(text,'<table width="100%" class="order-list" border="2">','</table>')
        #Removing unnecessary styling and HTML tags
        data = re.sub(r"state\d{2}", "state", data)
        data = re.sub(r"state\d", "state", data)
        branch = find_between(data,'<th>Branch</th><td>','</td><th>')
        replace = [' id="state"','<th>Registration Number</th>','<th>Student Name</th>','<th>Degree</th><td>'+course+'</td></tr><tr><th>Branch</th><td>'+branch+'</td><th>Semester</th><td>'+sem+'</td><th>Month & Year </th><td>'+month+'-'+year+'</td></tr><tr><td><b>Subject Code</b></td><td><b>Subject Name</b></td><th>Marks (Grade)</th><th>Result</th></tr>',' style="text-align:center;"']
        replace2 = {'<tr><td>':'','</td><td>':'\t','</td>':'\t','</tr>':''}
        data = replace_all(data,replace)
        data = replace_all_2(data,replace2)
        data = data.replace("\tPASSED",'')
        data = data.replace("\tFAILED",'')
        #Getting total and GPA
        data = data.replace("<b>Subject Code</b>	<b>Subject Name</b>	<th>Marks (Grade)</th><th>Result</th>",'')
        total = find_between(text,'Total :','<br>')
        gpa = find_between(text,'GPA   :','<br>')
        clmn = data.split('\t')
        name = ""
        j = 0
        data = ""
        #Removing subject name
        while(j < len(clmn)):
            if ((j == 0) or (j%3 != 0)):
                data += clmn[j]+"\t"
            if (j == 1):
                name = clmn[1]
            j += 1
        print "Fetched : " + str(i) + " - " + name
        #Completely fetched data
        result  += data+total+'\t'+gpa+"\n"
    except:
        print "Error"
    i += 1

#Writing result to file
with open("Result.txt", "w") as myfile:
            myfile.write(result)
            
print "\n\n\nFile exported as text (Result.txt) , Copy-paste it to Excel"
print "Script by Kishor V"
raw_input("Press any key to exit...")

