# Easy Google Docs for Java
This is a simple library that uses the Google Docs RESTful APIs to create,
edit and delete documents, principally focussed on Sheets.

Google Sheets are a rich and sophisticated alternative to MS Office 365 Excel and
are every bit a match in features and functionality.

The APIs for Google Docs is extensive, rooted in RESTful and wrapped with other language
libraries that are entirely auto-generated. This is the case for the Java SDK.
To say that it's hard to use the first time you encounter it, is a bit of an understatement.
This is a nice abstraction of the SDK that simplifies handling spreadsheets
and files in the Google Docs domain.

This library aggregates the classes and methods contained in the Google libraries into
something that is easy, coherent and efficient to use.

## Usage
To add the library to your project you will need to add GitHub packages
to your 

Maven
```
<dependency>
  <groupId>com.pivotal-solutions</groupId>
  <artifactId>ezgdocs4j</artifactId>
  <version>1.0.0</version>
</dependency>
```
Gradle
```
compile "com.pivotal-solutions:ezgdocs4j:1.0.0"
```

## Examples
The unit tests are your friend here but to get you started, here's a
simple example of creating a spreadsheet in the root folder, listing
the files in the root and deleting the spreadsheet;
```java
GoogleFile rootFolder = new GoogleFile();
String ssName = String.format("test-%d", System.currentTimeMillis());
GoogleSpreadsheet ss = GoogleSpreadsheet.create(ssName);
assertEquals(ss.getName(), ssName, "New spreadsheet is misnamed");
List<GoogleFile> files = rootFolder.list((dir, name) -> name.equals(ssName));
assertFalse(files.isEmpty(), "No new files found");
ss.delete();
```

## Google Credentials
To use Google Workspace APIs (Google Docs), you must have an access token for each request. The client
library will take care of this, but you need Credentials in the form of a JSON file. Credentials belong
to a Google Project which can be personal or organisation level. 

In all cases, any Google Docs that are created/edited must be accessible using these Credentials so if for
instance you set up a Service Account, that account must have access to the Google Docs in your domain.
You can do this by adding your Service Account name (email address from the Credentials file) to the Share 
names from the UI.
Here is a link to the Google help [docs](https://developers.google.com/workspace/guides/manage-credentials).

When you get the JSON file you will need to put it somewhere that the application can find it.
It will look for the file in the following places (in this order);
- System variable `google.credentials.filename` 
- Environment variable `SYSTEM_GOOGLE_CREDENTIALS_FILENAME`
- A file called `credentials.json` in a folder called `.google` in the users root (home) directory

Token management/refresh is all handled by the Google API Client.

## Development
The library is a maven project with no external dependencies over and above those libraries defined
in the POM file.

### Debugging Google Workspace APIs
As mentioned earlier, all the Google APIs are machine generated from their RESTful counterparts which makes
them a little terse. You will also quickly find that the documentation for Java is thin on the ground
and often refers to the RESTful equivalents which are expressed in JSON POST/PUTs

To figure out what your calls are actually doing, you can turn on logging.
Google Java SPIs use `JDK logging` which is managed by a file called `logging.properties`. You can set an alternative
on the command line as a system variable but the absence of this, the application will attempt to load it
from the resource path.

Change the `java.util.logging.ConsoleHandler.level` to `CONFIG or something finer grained and you will see the
RESTful request/responses that are being exchanged by the APIs.  Very useful when trying to figure out what
fields you're missing or got wrong.

### Google Batching and Request Rate Limits
The Google APIs have rate limits applied to them for writing that can be quite onerous.
A way to work around this is the 'batch' your requests - this is achieved by simply turning on batching using the
`batchStart()` method on a `GoogleSheet` instance. This stores the commands until you call `batchExecute()`.
This will send all your accumulated commands in a single POST request which helps not only performance, but also
the rate limits.

For example;

```java
// First of all, open the spreadsheet
GoogleSpreadsheet spreadsheet = new GoogleSpreadsheet(spreadSheetId);

// Get the results
CalculatedResults results = getServiceCostsForAllAccounts(referenceSpreadSheetId, range, excludeZero, statuses);

// Add the summary data
spreadsheet.clear();
GoogleSheet sheet = spreadsheet.addSheet("Summary", 0, false);
sheet.batchStart();
sheet.mergeCells("A1:L1",GoogleSheet.MergeType.MERGE_ALL);
sheet.formatCells("A1:L1").bold().fontSize(11).alignLeft().alignMiddle().alignLeft().apply();
sheet.appendValues("A2","Name","ID","Status","Owner","Service","Is Valid?","Type","Date","Month","Recurring","Cost","Adjusted Cost");
sheet.formatCells("A2:L2").backgroundColor(Color.GRAY).foregroundColor(Color.WHITE).bold().apply();
sheet.formatCells("J2:L2","H").alignRight().apply();
sheet.setNote("L2","This is a note about what this cell contains");

// Add the CSV values
sheet.appendValues("A3",results.getCostData());
sheet.formatCells("H3:H").date().numberPattern("d mmmm yy").apply();
sheet.formatCells("K2:L").currency().apply();

sheet.formatCells("H3:H","K2:L","A2:B").alignRight().apply();
sheet.formatCells("F2:G","J2:J").alignCenter().apply();
sheet.freezeRowsAndColumns(2, null);
sheet.autoResizeColumns("A");
sheet.setText("A1",String.format("Period from %s%s", range, excludeZero ?" (Excludes costs below 1 cent)":""));
sheet.batchExecute();
```

Note: - *A good way to debug a sheet, is to turn off the batching and let each command be executed on the server
as you step through the code in the IDE. Updates to the sheet are real-time.*

### Building the project
#### Prerequisites
- Java 1.8+
- Apache Maven

The application is purposely lightweight, it doesn't use a framework like Spring etc. This
is by design and is intended to keep the code base simple so that it is
easily understood and extendable with the minimum of knowledge.

The project uses a standard Maven POM
```
mvn clean package
```
This will create a jar e.g.
```
java -jar ezgdocs4j-1.0-SNAPSHOT.jar -h
```


