package sample.util;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import sample.Main;
import sample.model.StudentResponseMetadata;

import java.io.*;
import java.util.ArrayList;
import java.util.Arrays;

public class ExcelManager {
    private XSSFWorkbook xssfWorkbook;
    private File bookFile;
    private FormulaEvaluator formulaEvaluator;

    public ExcelManager() {
    }

    public XSSFWorkbook getBook(){
        return this.xssfWorkbook;
    }

    public void createNewAnswerBook(File newLocation) { // used for new assessments
        try {
            final String path = System.getProperty("user.dir") + "/excelTemplates/AgilExcelTemplate.xlsx";
            FileInputStream templateInStream = new FileInputStream(new File(path));
            this.xssfWorkbook = new XSSFWorkbook(templateInStream);
            FileOutputStream newBookStreamOut = new FileOutputStream(newLocation); // write to the file location given with the name given
            this.xssfWorkbook.write(newBookStreamOut);
            this.bookFile = newLocation;
            formulaEvaluator = xssfWorkbook.getCreationHelper().createFormulaEvaluator();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public String addToExistingBook(File chosenFile) {
        try {
            File oldValueOfBookFile = new File("."); // empty file
            if (bookFile != null) {
                oldValueOfBookFile = new File(bookFile.toURI()); // make copy just in case it is bad
            }
            bookFile = chosenFile;
            if (canSave()) {
                String path = chosenFile.getPath();
                FileInputStream templateInStream = new FileInputStream(new File(path));
                if(this.xssfWorkbook != null){
                    this.xssfWorkbook.close();
                }
                this.xssfWorkbook = new XSSFWorkbook(templateInStream);
                formulaEvaluator = xssfWorkbook.getCreationHelper().createFormulaEvaluator();
                return "Adding to: " + bookFile.getName();
            }
            else {
                Main.promptWarningOpenFile();
                if (oldValueOfBookFile.getName().length() >= 2) {
                    bookFile = oldValueOfBookFile; // restore old file value
                    return "Still adding to: " + bookFile.getName();
                }
                return chosenFile.getName() + " must be closed.";
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
            Main.promptConfirm();
            return chosenFile.getName() + " must be closed.";
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }
    public ArrayList<StudentResponseMetadata>[] getAllStudentMetadataInSheet(){
        ArrayList<StudentResponseMetadata>[] allStudentResponses = new ArrayList[5]; // sheets A - E
        String[] allKeyVersions = {"A","B","C","D","E"};
        int studentResponseOffsetColAmount = 4;
        for (int version = 0; version < allKeyVersions.length; version++) {// cycle sheets A - E
            Sheet currentSheet = this.xssfWorkbook.getSheet("Key " + allKeyVersions[version]); // hold the sheet needed to read key
            int numStudents = currentSheet.getRow(3).getLastCellNum();
            ArrayList<StudentResponseMetadata> currentMetaList = new ArrayList<>();
            for (int student = 0; student < numStudents; student++) { // cycle everyone in that row
                int zeroBasedOffsetAdjustedColNumber = studentResponseOffsetColAmount + student - 1;
                StudentResponseMetadata currentStudent = new StudentResponseMetadata();
                currentStudent.setKeyVersion(allKeyVersions[version]); // set the key version
                System.out.println("Student: " + zeroBasedOffsetAdjustedColNumber);
                for(int rowNum = 3; rowNum <= 5; rowNum++){
                    Row currRow = currentSheet.getRow(rowNum);
                    if(currRow == null){
                        currRow = currentSheet.createRow(rowNum);
                    }
                    Cell currCell = currRow.getCell(zeroBasedOffsetAdjustedColNumber);
                    if(currCell != null) {
                        try {
                            if (rowNum == 3) {//id
                                String id = currCell.getStringCellValue();
                                currentStudent.setId(id);
                            } else if (rowNum == 4) { // numCorrect
                                int numCorrect = (int) currCell.getNumericCellValue();
                                currentStudent.setNumCorrect(numCorrect);
                            } else { // percentCorrect, (rowNum == 5)
                                double percentCorrect = currCell.getNumericCellValue();
                                currentStudent.setPercentCorrect(percentCorrect);
                            }
                        }
                        catch(IllegalStateException ise){
                            currentStudent.setPercentCorrect(0);
                        }
                    }
                }
                if(currentStudent.getId() != null && !currentStudent.getId().equals("")) // not null or blank
                    currentMetaList.add(currentStudent);
            }
            allStudentResponses[version] = currentMetaList;
        }
        return allStudentResponses;
    }

    public String[][] alterKeysBasedOnKeysFromSheet(String[][] currentKeysStoredInApp){
        int[] indexesOfNullKeys = new int[5]; // these are the keys we want to read from
        for(int keyInd = 0; keyInd < 5; keyInd++){
            if(currentKeysStoredInApp[keyInd] != null &&
                    currentKeysStoredInApp[keyInd][0] != null){ // first answer is "O"
                indexesOfNullKeys[keyInd] = 1; // unknown at the moment
            }
            else { // null, so we need to overwrite
                indexesOfNullKeys[keyInd] = 0; // unknown at the moment
            }
        }
        String[][] allKeys = new String[5][50]; // keys A - E
        String[] allKeyVersions = {"A","B","C","D","E"};
        int offset = 8;
        for (int version = 0; version < allKeyVersions.length; version++) {// cycle keys A - E
            Arrays.fill(allKeys[version], "O");
            if(indexesOfNullKeys[version] == 0) {
                Sheet currentSheet = this.xssfWorkbook.getSheet("Key " + allKeyVersions[version]); // hold the sheet needed to read key
                String[] keyResponse = new String[50];
                for (int row = 8; row < 50 + offset; row++) { // cycle the 50 questions
                    Row r = currentSheet.getRow(row - 1); // row iterator, zero-based
                    keyResponse[row - offset] = r.getCell(2).getStringCellValue(); // val for particular question
                }
                allKeys[version] = keyResponse;
            }
        }
        return allKeys;
    }

    public void saveBook() {
        try {
            try {
                XSSFFormulaEvaluator.evaluateAllFormulaCells(xssfWorkbook); // so that it is updated before user looks at it
            } catch (IllegalArgumentException iae) { // silently don't evaluate
                iae.printStackTrace();
            }
            FileOutputStream newBookStreamOut = new FileOutputStream(bookFile); // write to the file location given with the name given
            xssfWorkbook.write(newBookStreamOut);
            newBookStreamOut.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void closeBook() {
        try {
            saveBook();
            xssfWorkbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public void addStudentResponsesToBook(String testVersion, String id, String[] responses) {
        Sheet currentSheet = this.xssfWorkbook.getSheet("Key " + testVersion);
        int rowOffset = 4; // beginning on row 8
        int rowOffsetForResponses = 8;
        Row idRow = currentSheet.getRow(rowOffset - 1); // row 8 [zero - based indexing]
        if (idRow == null) { // if the cell doesnt exist, create a new one in its place
            currentSheet.createRow(rowOffset - 1);
        }
        int colToWriteTo = findIdInRow(idRow, id);
        
        Cell nameCell = currentSheet.getRow(2).getCell(colToWriteTo);
        // add ID
        if (nameCell == null) {
            nameCell = currentSheet.getRow(2).createCell(colToWriteTo);
        }
        nameCell.setCellFormula("VLOOKUP(" + CellReference.convertNumToColString(colToWriteTo) + "4, 'ID Name Table'!$A:$B, 2, FALSE)");
        
        Cell idCell = currentSheet.getRow(3).getCell(colToWriteTo);
        // add ID
        if (idCell == null) {
            idCell = currentSheet.getRow(3).createCell(colToWriteTo);
        }
        idCell.setCellValue(id);
        // add raw responses
        for (int row = rowOffsetForResponses; row < rowOffsetForResponses + 50 &&
                (row <= rowOffsetForResponses || !responses[row - rowOffsetForResponses].equals("O")); row++) { // begins at C8, then goes for 50 rows
            Row currentRow = currentSheet.getRow(row - 1); // row 8 [zero - based indexing]
            if (currentRow == null) { // if the cell doesnt exist, create a new one in its place
                currentSheet.createRow(row);
            }
            Cell currentCell = currentRow.getCell(colToWriteTo); // last column

            //System.out.println("Row: " + row + " Col: " + colToWriteTo);
            if (currentCell == null) { // if the cell doesnt exist, create a new one in its place
                currentCell = currentRow.createCell(colToWriteTo);
            }
            // normal answer entry
            currentCell.setCellValue(responses[row - rowOffsetForResponses]);
        }
        // add true-false table
        int trueFalseRowNumStart = rowOffsetForResponses + 50 + 2; // line 60 in spreadsheet
        for (int row = trueFalseRowNumStart; row < trueFalseRowNumStart + 50; row++) { // begins at C8, then goes for 50 rows
            Row currentRow = currentSheet.getRow(row - 1); // row 8 [zero - based indexing]
            Cell currentCell = currentRow.getCell(colToWriteTo); // last column
            //System.out.println("Row: " + row + " Col: " + colToWriteTo);
            if (currentCell == null) { // if the cell doesnt exist, create a new one in its place
                currentRow.createCell(colToWriteTo);
                currentCell = currentRow.createCell(colToWriteTo);
            }
            // normal answer entry
            String columnLetter = CellReference.convertNumToColString(colToWriteTo); // " + (trueFalseRowNumStart - (50 + 2) + row) + "
            currentCell.setCellFormula(
                    "IF(AND(ISTEXT("
                            + columnLetter + (row - (50 + 2)) + ")," + columnLetter + (row - (50 + 2)) +
                            "=$C" + (row - (50 + 2)) + "),TRUE,FALSE)");
            formulaEvaluator.evaluateFormulaCell(currentCell);
        }
        //evaluate formulas
        Cell numCorrectCell = currentSheet.getRow(4).getCell(colToWriteTo);
        if (numCorrectCell == null) {
            numCorrectCell = currentSheet.getRow(4).createCell(colToWriteTo);
        }
        Cell percentCorrectCell = currentSheet.getRow(5).getCell(colToWriteTo);
        if (percentCorrectCell == null) {
            percentCorrectCell = currentSheet.getRow(5).createCell(colToWriteTo);
        }
        String columnLetter = CellReference.convertNumToColString(colToWriteTo);
        numCorrectCell.setCellFormula("COUNTIF(" + columnLetter + "60:" + columnLetter + "109, TRUE)");
        formulaEvaluator.evaluateFormulaCell(numCorrectCell);
        percentCorrectCell.setCellFormula("(" + columnLetter + "$5/$C$2)*100");
        formulaEvaluator.evaluateFormulaCell(percentCorrectCell); 
        formulaEvaluator.evaluateFormulaCell(nameCell);
    }

    public int findIdInRow(Row row, String searchVal) {
        int startCol = 3;
        int lastElem = row.getLastCellNum();
        //System.out.println("lastElem: " + lastElem);
        for (int i = startCol; i < lastElem; i++) {
            Cell currCell = row.getCell(i);
            currCell.setCellType(CellType.STRING);
            if (currCell.getStringCellValue().equals(searchVal)) {
                currCell.setCellType(CellType.STRING);
                return i;
            } else {
                currCell.setCellType(CellType.STRING);
            }
        }
        return lastElem;
    }

    // c8 is where the entries begin
    public void addKeyToBook(String version, String[] responses) {
        int rowOffset = 8; // beginning on row 8
        Sheet currentSheet = this.xssfWorkbook.getSheet("Key " + version);
        for (int row = rowOffset; row < rowOffset + 50 && !responses[row - rowOffset].equals("O"); row++) { // begins at C8, then goes for 50 rows
            Row currentRow = currentSheet.getRow(row - 1); // row 8 to start [zero - based indexing]
            Cell currentCell = currentRow.getCell(2); // column C, [zero-based indexing]
            currentCell.setCellValue(responses[row - rowOffset]);
        }
    }

    public boolean isAssigned() {
        return this.bookFile != null && this.xssfWorkbook != null;
    }

    public boolean canSave() {
        File fCopy = new File(bookFile.getAbsolutePath());
        return bookFile.renameTo(fCopy);
    }

    public void deleteColumn(Sheet sheet, int columnToDelete){
        for (int rId = 0; rId <= sheet.getLastRowNum(); rId++) {
            Row row = sheet.getRow(rId);
            for (int cID = columnToDelete; cID < row.getLastCellNum(); cID++) {
                Cell cOld = row.getCell(cID);
                if (cOld != null) {
                    row.removeCell(cOld);
                }
                Cell cNext = row.getCell(cID + 1);
                if (cNext != null) {
                    Cell cNew = row.createCell(cID, cNext.getCellType());
                    cloneCell(cNext, cNew);
                    sheet.setColumnWidth(cID, sheet.getColumnWidth(cID + 1));
                }
            }
        }
    }
    private void cloneCell( Cell cNew, Cell cOld ){
        cNew.setCellComment( cOld.getCellComment() );
        cNew.setCellStyle( cOld.getCellStyle() );

        switch ( cNew.getCellType() ){
            case BOOLEAN:{
                cNew.setCellValue( cOld.getBooleanCellValue() );
                break;
            }
            case NUMERIC:{
                cNew.setCellValue( cOld.getNumericCellValue() );
                break;
            }
            case STRING:{
                cNew.setCellValue( cOld.getStringCellValue() );
                break;
            }
            case ERROR:{
                cNew.setCellValue( cOld.getErrorCellValue() );
                break;
            }
            case FORMULA:{
                cNew.setCellFormula( cOld.getCellFormula() );
                break;
            }
        }

    }
}
