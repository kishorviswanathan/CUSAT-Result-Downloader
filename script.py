import urllib2
import re
import urllib


#Function definitions
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
print "\n\nOpen a result on the webpage and input values as per the website\n"



#Variable initialisation
courses = ["B.Tech","M.Tech"]
exams = ["Regular","Revaluation","Supplementary","Improvement"]
months = ["January","February","March","April","May","June","July","August","September","October","November","December"]
result = ""
branch = ""

#Getting user inputs
sem = raw_input("Semester : ")

#Loop until user enters correct input
while(1):
	course = int(input("\nDegree [ 1. B.Tech | 2. M.Tech ] (Integer) : "))
	if(course > 0 and course < 3):
		course = courses[course - 1]
		break	
	else:
		print "Invalid Entry"

while (1) :
	exam = int(input("\nResult type [1. Regular | 2. Revaluation | 3. Supplementary | 4. Improvement ] (Integer) : "))
	if(exam > 0 and exam < 5):
		exam = exams[exam - 1]
		break	
	else:
		print "Invalid Entry"

while (1) :
	month = int(input("\nMonth (Integer) :"))
	if(month > 0 and month < 13):
		month = months[month - 1]
		break
	else:
		print "Invalid Entry"

year = raw_input("\nYear : ")
i = int(input("\nStarting Reg. Number : "))
last = int(input("\nLast Reg. Number : "))
filename = raw_input("\nOutput Excel Sheet Name [Without extension] : ")
filename = filename+".xls"
url = "http://exam.cusat.ac.in/erp5/cusat/CUSAT-RESULT/Result_Declaration/display_sup_result?statuscheck=failed&deg_name="+course+"&semester="+urllib.quote(sem)+"&month="+month+"&year="+year+"&result_type="+exam+"&regno="
            


#Printing headings
print "\n\n\nReg. No".ljust(11) + "Name".ljust(30) + "Semester".ljust(15) + "Course".ljust(15) + "Classification"
print "-"*100 + "\n"

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

        #Getting Total and GPA
        data = data.replace("<b>Subject Code</b>	<b>Subject Name</b>	<th>Marks (Grade)</th><th>Result</th>",'')
        total = find_between(text,'Total :','<br>')
        gpa = find_between(text,'GPA   :','<br>')
        gpa = gpa + "0"
        tgpa = find_between(text,'CGPA&nbsp;:&nbsp;','</b><br>')
        tgpa = tgpa + "0"
        ennelum= find_between(text,'Classification&nbsp;:&nbsp;','</b><br>')
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
	
	#We don't have a result, check next
	if(name == "<b>Subject Name</b>" or name == ""):
            print str(i) + "   Result unavailable"
            i += 1
            continue

	#Justifying text for printing
        name = name.ljust(30)
        toprint = str(i) + "   " + name

	#Checking status
        if(float(gpa)>0):
            toprint = toprint + "PASSED".ljust(15)
        else:
            toprint = toprint + "FAILED".ljust(15)
        if (sem == '8'):
            if(float(tgpa)>0):
                toprint = toprint + "PASSED".ljust(15) +ennelum
            else:
                toprint = toprint + "FAILED".ljust(15)
	#Printing to screen
        print toprint

        result  += data+total+'\t'+gpa+ '\t' +tgpa + '\t' + ennelum +"\n"
    except Exception,e:
        print e

    i += 1

#Writing result to file
with open(filename, "w") as myfile:
	myfile.write(result)
            
print "\n\n\nFile exported as Excel Sheet : " + filename
print "\n\n\nScript by Kishor V"
raw_input("Press any key to exit...")
