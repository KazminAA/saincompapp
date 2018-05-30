import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.IOException;
import java.util.*;
import java.util.stream.Collectors;

public class ElFileParser {
    private Map<String, ParsedRow> resultMap;
    private int[] parsingColumn;

    private static ParsedRow applySum(ParsedRow r1, ParsedRow r2) {
        r1.setNumber(r1.getNumber() + r2.getNumber());
        return r1;
    }

    public Map<String, ParsedRow> getDifference(File firstFile, File seconFile, int[] parsingColumn) throws IOException, InvalidFormatException {
        resultMap = new TreeMap<>();
        this.parsingColumn = parsingColumn;
        List<ParsedRow> rows = readRowFromExcel(firstFile);
        rows.forEach(row -> {
            if (row.getSain() == null || row.getSain().isEmpty()) {
                resultMap.merge(row.getName() + " " + row.getOperation(), row, ElFileParser::applySum);
            } else {
                resultMap.merge(row.getSain() + " " + row.getOperation(), row, ElFileParser::applySum);
            }
        });
        resultMap.values().forEach(parsedRow -> parsedRow.setNumber(parsedRow.getNumber() * -1));
        rows = readRowFromExcel(seconFile);
        rows.forEach(row -> {
            if (row.getSain() == null || row.getSain().isEmpty()) {
                resultMap.merge(row.getName() + " " + row.getOperation(), row, ElFileParser::applySum);
            } else {
                resultMap.merge(row.getSain() + " " + row.getOperation(), row, ElFileParser::applySum);
            }
        });
        return resultMap.values().stream().filter(row -> row.getNumber() != 0).collect(
                Collectors.toMap(parsedRow -> {
                    if (parsedRow.getSain() == null || parsedRow.getSain().isEmpty()) {
                        return parsedRow.getName() + " " + parsedRow.getOperation();
                    } else {
                        return parsedRow.getSain() + " " + parsedRow.getOperation();
                    }
                }, parsedRow -> parsedRow)
        );
    }

    private List<ParsedRow> readRowFromExcel(File file) throws IOException, InvalidFormatException {
        Workbook workbook = WorkbookFactory.create(file);
        Sheet sheet = workbook.getSheetAt(0);
        List<ParsedRow> result = new ArrayList<>();
        DataFormatter dataFormatter = new DataFormatter();
        Iterator<Row> rowIterator = sheet.rowIterator();
        for (int i = 0; i < parsingColumn[4]; i++) {
            if (rowIterator.hasNext()) {
                rowIterator.next();
            }
        }
        while (rowIterator.hasNext()) {
            Row elRow = rowIterator.next();
            ParsedRow row = new ParsedRow();
            String cell2 = dataFormatter.formatCellValue(elRow.getCell(parsingColumn[2]));
            if (cell2 == null || !cell2.matches("\\d+")) {
                continue;
            }
            row.setSain(dataFormatter.formatCellValue(elRow.getCell(parsingColumn[0])));
            row.setName(dataFormatter.formatCellValue(elRow.getCell(parsingColumn[1])));
            row.setNumber(Integer.valueOf(cell2));
            row.setNotes(dataFormatter.formatCellValue(elRow.getCell(parsingColumn[3])));
            row.setOperation(dataFormatter.formatCellValue(elRow.getCell(parsingColumn[5])));
            result.add(row);
        }
        return result;
    }

}
