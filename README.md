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

##  Script Assertion 
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

## Get the outcome of an assertion (in this case Script Assertion)
```groovy
// get outcome of the test
def testOutCome =context.testCase.getTestStepByName('<stepName>').getAssertionByName('Script Assertion').status.toString()
log.info(testOutCome)
```

## Add both Request & Response to the spreadsheet
```groovy
//create a new cell
def cellZero = row.createCell(0)
def cellOne = row.createCell(1)
// Set the cell value
def xmlRequest = context.expand('${Eligibility270#Request}')
log.info(xmlRequest)
cellZero.setCellValue(xmlRequest)
def xmlResponse = context.expand('${Eligibility270#Response}')
cellOne.setCellValue(xmlResponse)
```
## Add column headers to excel
```groovy
//get last row number
def lastRowNum = sheet.getLastRowNum()
log.info("last Row " + lastRowNum)
def row = sheet.createRow(0)
if(lastRowNum==-1){
row.createCell(0).setCellValue("RequestXml")
row.createCell(1).setCellValue("ResponseXml")
lastRowNum++
}
//create a new row
row =sheet.createRow(lastRowNum+1)

//create a new cell
def cellZero = row.createCell(0)
def cellOne = row.createCell(1)
// Set the cell value
def xmlRequest = context.expand('${NullEdiHeader#Request}')
log.info(xmlRequest)
cellZero.setCellValue(xmlRequest)
def xmlResponse = context.expand('${NullEdiHeader#Response}')
cellOne.setCellValue(xmlResponse)

// Write the workbook to a file
def file = new File(filePath.toString())
def outputStream = new FileOutputStream(file)
workbook.write(outputStream)
outputStream.close()
log.info "Response wrote to  : $file"
```
## Conditional Go To
https://support.smartbear.com/readyapi/docs/functional/steps/goto.html?sbsearch=Conditional

## Generate Time Stamps
```groovy
import java.text.SimpleDateFormat
import java.util.Calendar

//Get current Time stamp
def currentTimeStamp = new Date()
def sdf1 = new SimpleDateFormat("yyMMdd") //format to yymmdd
def formattedCurrentTimeStamp1 = sdf1.format(currentTimeStamp)

def sdf2 = new SimpleDateFormat("yyyyMMdd") //format to yymmdd
def formattedCurrentTimeStamp2 = sdf2.format(currentTimeStamp)

log.info("Current Time Stamp1:  $formattedCurrentTimeStamp1")
log.info("Current Time Stamp2:  $formattedCurrentTimeStamp2")

//get timestamp minus one month
def cal = Calendar.getInstance()
cal.add(Calendar.MONTH, -1)
def minusOneMonthTimeStamp = cal.time

def formattedMinusOneMonthTimeStamp = sdf2.format(minusOneMonthTimeStamp)

log.info("Minus One Month Time Stamp :  $formattedMinusOneMonthTimeStamp")

//set properties in properties test step
def propertiesTestStep = testRunner.testCase.getTestStepByName("Properties")
propertiesTestStep.setPropertyValue("DATE1", formattedCurrentTimeStamp1)
propertiesTestStep.setPropertyValue("DATE2", formattedCurrentTimeStamp2)
propertiesTestStep.setPropertyValue("SERVICEDATE", formattedMinusOneMonthTimeStamp)
```
