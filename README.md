# SOAPUIForMe
## Groovy Script to Add Response to Excel
Add latest poi.jar and poi-ooxml to bin/ext directory for soapui
```groovy
import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.apache.poi.xssf.usermodel.XSSFCellStyle
import org.apache.poi.xssf.usermodel.XSSFFont
import org.apache.poi.xssf.usermodel.XSSFColor
import java.nio.file.*
import java.text.SimpleDateFormat
import java.util.Date

//get the project path dynamically
def projectFile = context.testCase.testSuite.project.getWorkspace().getPath()
def projectPath = new File(projectFile).parent
//append the ile name to project path
def dateFormat = new SimpleDateFormat("yyyyMMdd_HHmmss")
def timeStamp = dateFormat.format(new Date())
def fileName="output_${timeStamp}.xlsx"
def filePath = Paths.get(projectPath, fileName)
log.info filePath.toString()

def workbook = new XSSFWorkbook()
if (Files.exists(filePath)){
 workbook = new XSSFWorkbook(Files.newInputStream(filePath))

}else{
// Create a new workbook
 workbook = new XSSFWorkbook()
 workbook.createSheet("sheet1")
 Files.createDirectories(filePath.getParent())
 Files.createFile(filePath)
 workbook.write(Files.newOutputStream(filePath))
}
// get first sheet
def sheet = workbook.getSheetAt(0)



//get last row number
def lastRowNum = sheet.getLastRowNum()

//create a new row
def row =sheet.createRow(lastRowNum+1)

//create a new cell
def cell = row.createCell(0)
// Set the cell value
def xmlResponse = context.expand('${TeamNames#Response}')
cell.setCellValue(xmlResponse)

// Write the workbook to a file
def file = new File(filePath.toString())
def outputStream = new FileOutputStream(file)
workbook.write(outputStream)
outputStream.close()
log.info "Response wrote to  : $file"
```

##  Sript Assertion 
```groovy
import com.eviware.soapui.support.XmlHolder
def response = context.expand('${<RequestStep>#Response}')
def parseXmlResponse = new XmlHolder(response)
log.info parseXmlResponse.getNodeValue('<xpath>')
def memberId= context.expand('${MemberIDs#MemberID}')
def transactionType = '<Expected data>'
assert  parseXmlResponse.getNodeValue('<xpath>').contains('<text>'):"<Error message when not matching>"
assert  parseXmlResponse.getNodeValue('<xpath>').contains('<text to be validated>'):"<Error message when not matching>"
```
