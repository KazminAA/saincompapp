import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.WorkbookUtil;

import java.io.File;
import java.io.IOException;
import java.util.*;

public class ElFileParser {
    private Map<String, ParsedRow> resultMap;
    private int[] parsingColumn;

    public Map<String, ParsedRow> getDifference(File firstFile, File seconFile, int[] parsingColumn) throws IOException, InvalidFormatException {
        resultMap = new HashMap<>();
        this.parsingColumn = parsingColumn;
        List<ParsedRow> rows = readRowFromExcel(firstFile);
        rows.stream().forEach(row -> resultMap.merge(
                row.getSain(),
                row,
                (r1, r2) -> {
                    r1.setNumber(r1.getNumber() + r2.getNumber());
                    return r1;
                }
        ));
        rows = readRowFromExcel(seconFile);
        rows.stream().forEach(row -> resultMap.merge(
                row.getSain(),
                row,
                (r1, r2) -> {
                    r2.setNumber(r2.getNumber() - r1.getNumber());
                    return r2;
                }
        ));
        return resultMap;
    }

    private List<ParsedRow> readRowFromExcel(File file) throws IOException, InvalidFormatException {
        Workbook workbook = WorkbookFactory.create(file);
        Sheet sheet = workbook.getSheetAt(0);
        List<ParsedRow> result = new ArrayList<>();
        DataFormatter dataFormatter = new DataFormatter();
        Iterator<Row> rowIterator = sheet.rowIterator();
        while (rowIterator.hasNext()) {
            Row elRow = rowIterator.next();
            ParsedRow row = new ParsedRow();
            row.setSain(dataFormatter.formatCellValue(elRow.getCell(parsingColumn[0])));
            row.setName(dataFormatter.formatCellValue(elRow.getCell(parsingColumn[1])));
            row.setNumber(Integer.valueOf(dataFormatter.formatCellValue(elRow.getCell(parsingColumn[2]))));
            result.add(row);
        }
        return result;
    }

}
