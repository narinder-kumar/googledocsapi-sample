package com.googledocs.sample;

import java.io.IOException;
import java.net.URL;
import java.util.List;
import java.util.Scanner;

import com.google.gdata.client.spreadsheet.CellQuery;
import com.google.gdata.client.spreadsheet.FeedURLFactory;
import com.google.gdata.client.spreadsheet.SpreadsheetService;
import com.google.gdata.data.spreadsheet.CellEntry;
import com.google.gdata.data.spreadsheet.CellFeed;
import com.google.gdata.data.spreadsheet.SpreadsheetEntry;
import com.google.gdata.data.spreadsheet.SpreadsheetFeed;
import com.google.gdata.data.spreadsheet.WorksheetEntry;
import com.google.gdata.util.AuthenticationException;
import com.google.gdata.util.ServiceException;

public class SpreadSheetDemo {

    private SpreadsheetService service;

    private FeedURLFactory factory;

    public static void main(String[] args) throws IOException, ServiceException {
        System.out.println("The program scans Google Docs to find any Spreadsheets created/shared by you");
        System.out.println("It also lets you find various details of your chosen Spreadsheet like worksheet, all cells within a worksheet or search for individual cells");
        System.out.println("You can exit the program by entering -1 when prompted for any input");
        
        String username = getUserInputAsString("Enter your Google User Email Id");
        String password = getUserInputAsString("Enter your password (****It will be printed on Console***)");
        SpreadSheetDemo spreadSheetDemo = new SpreadSheetDemo(username, password);
        spreadSheetDemo.showAllSpreadSheets();

        int desiredSpreadsheetIndex = spreadSheetDemo.getUserInputAsInteger("Enter Index you want further details (0, 1, 2 etc)");
        SpreadsheetEntry spreadsheet = spreadSheetDemo.getSpreadSheetFromUserInput(desiredSpreadsheetIndex);
        spreadSheetDemo.showDetailsAboutSpreadSheet(spreadsheet);

        int desiredWorksheetIndex = spreadSheetDemo.getUserInputAsInteger("Enter Index you want further details (0, 1, 2 etc)");
        WorksheetEntry worksheetEntry = spreadSheetDemo.getWorksheetFromUserInput(spreadsheet, desiredWorksheetIndex);
        spreadSheetDemo.showDetailsAboutWorksheet(worksheetEntry);

        getUserInputAsString("Would like to search for a given cell, Enter -1 to exit");
        int desiredStartRow = spreadSheetDemo.getUserInputAsInteger("Enter starting row number");
        int desiredEndRow = spreadSheetDemo.getUserInputAsInteger("Enter ending row number");
        int desiredStartColumn = spreadSheetDemo.getUserInputAsInteger("Enter starting column number");
        int desiredEndColumn = spreadSheetDemo.getUserInputAsInteger("Enter ending column number");
        spreadSheetDemo.showCellDetailsWithInRange(worksheetEntry, desiredStartRow, desiredEndRow, desiredStartColumn, desiredEndColumn);
        System.out.println("Exiting the program...");
        System.exit(0);
    }

    public SpreadSheetDemo(String username, String password) throws AuthenticationException {
        this.service = new SpreadsheetService("SpreadSheet-Demo");
        this.factory = FeedURLFactory.getDefault();        
        System.out.println("Authenticating...");
        this.service.setUserCredentials(username, password);
        System.out.println("Successfully authenticated");
    }
    
    public void showAllSpreadSheets() throws IOException, ServiceException {
        SpreadsheetFeed feed = service.getFeed(factory.getSpreadsheetsFeedUrl(), SpreadsheetFeed.class);
        List<SpreadsheetEntry> spreadsheets = feed.getEntries();
        System.out.println("Total number of SpreadSheets found : " + spreadsheets.size());
        for (int i = 0; i < spreadsheets.size(); ++i) {
            System.out.println("("+i+") : "+spreadsheets.get(i).getTitle().getPlainText());
        }
    }

    private void showDetailsAboutSpreadSheet(SpreadsheetEntry spreadsheet) throws IOException, ServiceException {
        System.out.println("SpreadSheet Title : "+spreadsheet.getTitle().getPlainText());
        List<WorksheetEntry> worksheets = spreadsheet.getWorksheets();
        for (int i = 0; i < worksheets.size(); ++i) {
            WorksheetEntry worksheetEntry = worksheets.get(i);
            System.out.println("("+i+") Worksheet Title : "+worksheetEntry.getTitle().getPlainText()+", num of Rows : "+worksheetEntry.getRowCount()+", num of Columns : "+worksheetEntry.getColCount());
        }
    }

    private void showDetailsAboutWorksheet(WorksheetEntry worksheetEntry) throws IOException, ServiceException {
        URL cellFeedUrl = worksheetEntry.getCellFeedUrl();
        CellFeed feed = service.getFeed(cellFeedUrl, CellFeed.class);
        System.out.println("Showing all cells of the Worksheet");
        for (CellEntry cell : feed.getEntries()) {
            System.out.println("\tTitle : "+cell.getTitle().getPlainText()+"\tAddress : "+cell.getId().substring(cell.getId().lastIndexOf('/') + 1));
            System.out.println("\tFormula : "+cell.getCell().getInputValue()+"\tValue : "+cell.getCell().getValue());
//            System.out.println("Calculated value : "+cell.getCell().getNumericValue());
        }        
    }

    private void showCellDetailsWithInRange(WorksheetEntry worksheetEntry, int startRow, int endRow, int startColumn, int endColumn) throws IOException, ServiceException {
        System.out.println("Showing Cell details within rows from "+startRow+ " to "+endRow+" and within columns from "+startColumn+" to "+endColumn);
        URL cellFeedUrl = worksheetEntry.getCellFeedUrl();
        CellQuery query = new CellQuery(cellFeedUrl);
        query.setMinimumRow(startRow);
        query.setMaximumRow(endRow);
        query.setMinimumCol(startColumn);
        query.setMaximumCol(endColumn);
        
        CellFeed queryFeedUrl = service.query(query, CellFeed.class);
        for (CellEntry cell : queryFeedUrl.getEntries()) {
            System.out.println("\tTitle : "+cell.getTitle().getPlainText()+"\tAddress : "+cell.getId().substring(cell.getId().lastIndexOf('/') + 1));
            System.out.println("\tFormula : "+cell.getCell().getInputValue()+"\tValue : "+cell.getCell().getValue());            
        }
    }
    
    private SpreadsheetEntry getSpreadSheetFromUserInput(int spreadsheetIndex) throws IOException, ServiceException {
        SpreadsheetFeed feed = service.getFeed(factory.getSpreadsheetsFeedUrl(), SpreadsheetFeed.class);
        SpreadsheetEntry spreadsheet = feed.getEntries().get(spreadsheetIndex);
        return spreadsheet;
    }

    private WorksheetEntry getWorksheetFromUserInput(SpreadsheetEntry spreadsheet, int worksheetIndex) throws IOException, ServiceException {
        WorksheetEntry worksheetEntry = spreadsheet.getWorksheets().get(worksheetIndex);
        return worksheetEntry;
    }
    
    private int getUserInputAsInteger(String inputQuestion) {
        Scanner scanner = new Scanner (System.in);
        System.out.print(inputQuestion + " : ");  
        String userInputString = scanner.next();
        int userInput = Integer.parseInt(userInputString);
        if (userInput == -1) {
            System.out.println("Exiting the program");
            System.exit(0);
        }
        return userInput;
    }

    private static String getUserInputAsString(String inputQuestion) {
        Scanner scanner = new Scanner (System.in);
        System.out.print(inputQuestion + " : ");  
        String userInputString = scanner.next();
        if ("-1".equals(userInputString)) {
            System.out.println("Exiting the program");
            System.exit(0);
        }
        return userInputString;
    }
}
