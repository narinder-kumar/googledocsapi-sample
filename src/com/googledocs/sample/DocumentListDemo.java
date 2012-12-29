package com.googledocs.sample;

import java.io.IOException;
import java.net.URL;
import java.util.Scanner;

import com.google.gdata.client.docs.DocsService;
import com.google.gdata.data.DateTime;
import com.google.gdata.data.Link;
import com.google.gdata.data.PlainTextConstruct;
import com.google.gdata.data.docs.DocumentEntry;
import com.google.gdata.data.docs.DocumentListEntry;
import com.google.gdata.data.docs.DocumentListFeed;
import com.google.gdata.data.docs.FolderEntry;
import com.google.gdata.data.docs.PresentationEntry;
import com.google.gdata.data.docs.SpreadsheetEntry;
import com.google.gdata.util.AuthenticationException;
import com.google.gdata.util.ServiceException;

public class DocumentListDemo {

    private static final String URL_SPREADSHEETS = "https://docs.google.com/feeds/default/private/full/-/spreadsheet";

    private static final String URL_ALL_DOCUMENTS = "https://docs.google.com/feeds/default/private/full/-/document";

    private static final String URL_ALL_FILES = "https://docs.google.com/feeds/default/private/full";

    private DocsService service;


    public static void main(String[] args) throws IOException, ServiceException {
        System.out.println("This program allows multiple CRUD operations on user's Google Documents");
        printUsage();
        String username = getUserInputAsString("Enter your Google User Email Id");
        String password = getUserInputAsString("Enter your password (****It will be printed on Console***)");
        DocumentListDemo documentListDemo = new DocumentListDemo(username, password);
        while (documentListDemo.parseInputAndProcessCommand(documentListDemo)) {
            //Continue till exit option is received
        }
    }

    private boolean parseInputAndProcessCommand(DocumentListDemo documentListDemo) throws IOException, ServiceException {
        String userInput = getUserInputAsString("Your Option");
        if ("help".equalsIgnoreCase(userInput)) {
            printUsage();
        } else if ("listAll".equalsIgnoreCase(userInput)) {
            documentListDemo.showElementDetails(URL_ALL_FILES);
        } else if ("listDocuments".equalsIgnoreCase(userInput)) {
            documentListDemo.showElementDetails(URL_ALL_DOCUMENTS);
        } else if ("listSpreadSheets".equalsIgnoreCase(userInput)) {
            documentListDemo.showElementDetails(URL_SPREADSHEETS);
        } else if ("create".equalsIgnoreCase(userInput)) {
            documentListDemo.createElement(URL_ALL_FILES);
        } else if ("trash".equalsIgnoreCase(userInput)) {
            documentListDemo.trashElement(URL_ALL_FILES);                        
        } else {
            System.out.println("Unknown command. Type 'help' for a list of commands.");
        }
        return true;

    }

    private static void printUsage() {
        System.out.println("List of Possible Options : ");
        System.out.println("listAll : Lists all files stored in your Google drive");
        System.out.println("listDocuments : Lists all documents stored in your Google drive");
        System.out.println("listSpreadSheets : Lists all SpreadSheets stored in your Google drive");
        System.out.println("create : Creates a new Document/SpreadSheet/Presentation/Folder in your Google Drive");
        System.out.println("trash : Trashes a file stored in your Google drive");
        System.out.println("-1 : Exit the program");
    }

    public DocumentListDemo(String username, String password) throws AuthenticationException {
        this.service = new DocsService("DocumentList-Demo");
        System.out.println("Authenticating...");
        this.service.setUserCredentials(username, password);
        System.out.println("Successfully authenticated");
    }

    private void showElementDetails(String urlParameter) throws IOException, ServiceException {
        URL url = new URL(urlParameter);
        DocumentListFeed documentListFeed = service.getFeed(url, DocumentListFeed.class);
        for (DocumentListEntry entry : documentListFeed.getEntries()) {
            printBriefSummary(entry);
        }
    }

    private void createElement(String urlParameter) throws IOException, ServiceException {
        String type = getUserInputAsString("What would you like to create : document/presentation/spreadsheet/folder");
        String title = getUserInputAsString("Name ");
        if (type == null || title == null) {
            throw new RuntimeException("null title or type");            
        }
        DocumentListEntry newEntry = null;
        if (type.equals("document")) {
          newEntry = new DocumentEntry();
        } else if (type.equals("presentation")) {
          newEntry = new PresentationEntry();
        } else if (type.equals("spreadsheet")) {
          newEntry = new SpreadsheetEntry();
        } else if (type.equals("folder")) {
          newEntry = new FolderEntry();
        }

        URL url = new URL(urlParameter);
        newEntry.setTitle(new PlainTextConstruct(title));
        DocumentListEntry newDocumentListEntry = service.insert(url, newEntry);
        if (newDocumentListEntry != null) {
            printBriefSummary(newDocumentListEntry);
        }
    }

    private void trashElement(String urlParameter) throws IOException, ServiceException {
        String resourceId = getUserInputAsString("Enter ResourceId of Element to trash");
        URL elementToDeleteURL = new URL(URL_ALL_FILES + "/" + resourceId);
        DocumentListEntry documentListEntry = service.getEntry(elementToDeleteURL, DocumentListEntry.class);

        URL trashElementUrl = new URL(URL_ALL_FILES+"/"+resourceId);
        service.delete(trashElementUrl, documentListEntry.getEtag());
        System.out.println("Successfully trashed element");
    }
    
    private void printBriefSummary(DocumentListEntry doc) {
        StringBuffer output = new StringBuffer();

        output.append(" -- " + doc.getTitle().getPlainText() + " ");
        if (!doc.getParentLinks().isEmpty()) {
            for (Link link : doc.getParentLinks()) {
                output.append("[" + link.getTitle() + "] ");
            }
        }
        output.append(doc.getResourceId());

        System.out.println(output);
    }

    private static String getUserInputAsString(String inputQuestion) {
        Scanner scanner = new Scanner(System.in);
        System.out.print(inputQuestion + " : ");
        String userInputString = scanner.next();
        if ("-1".equals(userInputString)) {
            System.out.println("Exiting the program");
            System.exit(0);
        }
        return userInputString;
    }

    private void printDetailedSummary(DocumentListEntry doc) {
        String resourceId = doc.getResourceId();
        String docType = resourceId.substring(0, resourceId.lastIndexOf(':'));

        System.out.println("'" + doc.getTitle().getPlainText() + "' (" + docType + ")");
        System.out.println("  link to Google Docs: " + doc.getHtmlLink().getHref());
        System.out.println("  resource id: " + resourceId);

        // print the timestamp the document was last viewed
        DateTime lastViewed = doc.getLastViewed();
        if (lastViewed != null) {
            System.out.println("  last viewed: " + lastViewed.toString());
        }

        // print other useful metadata
        System.out.println("  last updated: " + doc.getUpdated().toString());
        System.out.println("  viewed by user? " + doc.isViewed());
        System.out.println("  writersCanInvite? " + doc.isWritersCanInvite().toString());
        System.out.println("  hidden? " + doc.isHidden());
        System.out.println("  starrred? " + doc.isStarred());
        System.out.println();
    }


}
