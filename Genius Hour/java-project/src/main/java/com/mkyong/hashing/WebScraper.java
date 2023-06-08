package com.mkyong.hashing;

import java.io.BufferedReader;
import java.io.InputStreamReader;
import java.net.URL;
import java.net.URLConnection;
import java.io.FileOutputStream;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.*;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import org.apache.commons.lang3.StringEscapeUtils;
import org.jsoup.nodes.Entities;

public class WebScraper {
  public static void main(String[] args) throws Exception {
    Workbook workbook = new XSSFWorkbook();

    //create dedicated sheet and header row for all of the stats to be in one place
    Sheet allStatsSheet = workbook.createSheet("All Stats");
    Row allStatsHeader = allStatsSheet.createRow(0);
    
    
    //vvvvvvvvvvvvvvvSTART POSTS SCRIPTvvvvvvvvvvvvvvv
    //#region
    Sheet postsSheet = workbook.createSheet("Player Posts"); //create dedicated page to posts hit stat
    Row headerRow = postsSheet.createRow(0); //create header row for posts stat sheet
    Cell postsNameHeader = headerRow.createCell(0);
    postsNameHeader.setCellValue("Name");
    Cell postsPostHeader = headerRow.createCell(1);
    postsPostHeader.setCellValue("Posts + Crossbars Hit");
    int rowInd = 0;
    
    //connect to NHL stats to see how many records they have
    URL initUrl = new URL(
        "https://api.nhle.com/stats/rest/en/skater/realtime?isAggregate=true&isGame=false&sort=%5B%7B%22property%22:%22missedShotGoalpost%22,%22direction%22:%22DESC%22%7D,%7B%22property%22:%22playerId%22,%22direction%22:%22ASC%22%7D%5D&start=0&limit=100&factCayenneExp=gamesPlayed%3E=1&cayenneExp=gameTypeId=2%20and%20seasonId%3C=20222023%20and%20seasonId%3E=20222023");
    URLConnection initConnection = initUrl.openConnection();
    BufferedReader in = new BufferedReader(new InputStreamReader(initConnection.getInputStream()));

    // Read the contents of the web page line by line and store them in a StringBuilder
    StringBuilder content = new StringBuilder();
    String line;
    while ((line = in.readLine()) != null) {
      content.append(line);
      content.append(System.lineSeparator());
    }
    JSONObject initJSON = new JSONObject(content.toString());
    int limit = initJSON.getInt("total");
    int pgTotal = (int) (limit / 100)+1;
    System.out.println("Post Pages: " + pgTotal);
    // Close the BufferedReader
    in.close();

    //loop thorugh NHL's pages of JSONs and collect data from each record within the JSON
    for(int i = 0; i<pgTotal*100; i+=100){
      URL url = new URL(
          "https://api.nhle.com/stats/rest/en/skater/realtime?isAggregate=true&isGame=false&sort=%5B%7B%22property%22:%22missedShotGoalpost%22,%22direction%22:%22DESC%22%7D,%7B%22property%22:%22playerId%22,%22direction%22:%22ASC%22%7D%5D&start="+i+"&limit=100&factCayenneExp=gamesPlayed%3E=1&cayenneExp=gameTypeId=2%20and%20seasonId%3C=20222023%20and%20seasonId%3E=20222023");

      // Use a BufferedReader to read the contents of the web page
      URLConnection connection = url.openConnection();
      BufferedReader in1 = new BufferedReader(new InputStreamReader(connection.getInputStream()));

      // Read the contents of the web page line by line and store them in a StringBuilder
      StringBuilder content1 = new StringBuilder();
      String line1;
      while ((line1 = in1.readLine()) != null) {
        content1.append(line1);
        content1.append(System.lineSeparator());
      }
      JSONObject JSON = new JSONObject(content1.toString());
      JSONArray data = JSON.getJSONArray("data");

      //loop through the each profile, collecting individual stats
      for(int z = 0; z<data.length(); z++)
      {
        String name = data.getJSONObject(z).getString("skaterFullName");
        int posts = data.getJSONObject(z).optInt("missedShotGoalpost") + data.getJSONObject(z).optInt("missedShotCrossbar");
        //add stats to spreadsheet
        Row row = postsSheet.createRow(z+rowInd+1);
        Cell nameCell = row.createCell(0);
        nameCell.setCellValue(name);
        Cell postCell = row.createCell(1);
        postCell.setCellValue(posts);      
        
      }
      rowInd+=100;
      in1.close();
      
     }
     System.out.println("Posts done");

     //#endregion
     //^^^^^^END POSTS SCRIPT^^^^^^


//=============================================================================================


//vvvvvvSTART SHOTS, GOALS SCRIPTvvvvvv
     //#region
     Sheet goalSheet = workbook.createSheet("Player Goals"); //create specific sheet for the shot and goal totals 
     Row goalHeaderRow = goalSheet.createRow(0); //create header row and headers for the shot and goal sheet
     Cell goalsNameHeader = goalHeaderRow.createCell(0);
     goalsNameHeader.setCellValue("Name");
     Cell shotsHeader = goalHeaderRow.createCell(1);
     shotsHeader.setCellValue("Shots");
     Cell goalsHeader = goalHeaderRow.createCell(2);
     goalsHeader.setCellValue("Goals");
     int rowInd1 = 0;

     //connect to first page of records to determine total number of records
     URL initGoalUrl = new URL(
         "https://api.nhle.com/stats/rest/en/skater/summary?isAggregate=true&isGame=false&sort=%5B%7B%22property%22:%22shots%22,%22direction%22:%22DESC%22%7D,%7B%22property%22:%22playerId%22,%22direction%22:%22ASC%22%7D%5D&start=0&limit=100&factCayenneExp=gamesPlayed%3E=1&cayenneExp=gameTypeId=2%20and%20seasonId%3C=20222023%20and%20seasonId%3E=20222023");
     URLConnection initGoalConnection = initGoalUrl.openConnection();
     BufferedReader goalIn = new BufferedReader(new InputStreamReader(initGoalConnection.getInputStream()));
     // Read the contents of the web page line by line and store them in a StringBuilder
     StringBuilder goalContent = new StringBuilder();
     String ln;
     while ((ln = goalIn.readLine()) != null) {
       goalContent.append(ln);
       goalContent.append(System.lineSeparator());
     }
     JSONObject goalInitJSON = new JSONObject(goalContent.toString());
     int goalLimit = goalInitJSON.getInt("total");
     int goalpgTotal = (int) (goalLimit / 100)+1;
     System.out.println("Goal Pages: " + goalpgTotal); //determine total pages (how many times to loop)
 
     // Close the BufferedReader
     in.close();
 
    //loop through number of pages from NHL
     for(int i = 0; i<goalpgTotal*100; i+=100){
       
      URL url = new URL(
           "https://api.nhle.com/stats/rest/en/skater/summary?isAggregate=true&isGame=false&sort=%5B%7B%22property%22:%22shots%22,%22direction%22:%22DESC%22%7D,%7B%22property%22:%22playerId%22,%22direction%22:%22ASC%22%7D%5D&start="+i+"&limit=100&factCayenneExp=gamesPlayed%3E=1&cayenneExp=gameTypeId=2%20and%20seasonId%3C=20222023%20and%20seasonId%3E=20222023");
 
       // Use a BufferedReader to read the contents of the web page
       URLConnection connection = url.openConnection();
       BufferedReader in1 = new BufferedReader(new InputStreamReader(connection.getInputStream()));
 
       // Read the contents of the web page line by line and store them in a StringBuilder
       StringBuilder content1 = new StringBuilder();
       String line1;
       while ((line1 = in1.readLine()) != null) {
         content1.append(line1);
         content1.append(System.lineSeparator());
       }
       JSONObject JSON = new JSONObject(content1.toString());
       JSONArray data = JSON.getJSONArray("data");

       //loop through individual players on each page and collect stats
       for(int z = 0; z<data.length(); z++)
       {
         String name = data.getJSONObject(z).getString("skaterFullName");
         int shots = data.getJSONObject(z).optInt("shots");
         int goals = data.getJSONObject(z).optInt("goals");

         //add stats to the spreadsheet
         Row row = goalSheet.createRow(z+rowInd1+1);
         Cell nameCell = row.createCell(0);
         nameCell.setCellValue(name);
         Cell shotsCell = row.createCell(1);
         shotsCell.setCellValue(shots);  
         Cell goalCell = row.createCell(2);
         goalCell.setCellValue(goals);        
       }
       rowInd1+=100;

       in1.close();
       
      }
     System.out.println("Goals Done");
//#endregion
//^^^^^^END SHOTS, GOALS SCRIPT^^^^^^



//vvvvvvvvvvvvvvvvvvvvv START SALARY SCRIPT vvvvvvvvvvvvvvvvvvvv
//#region

CellStyle dollarStyle = workbook.createCellStyle();
CreationHelper creationHelper = workbook.getCreationHelper();
dollarStyle.setDataFormat(creationHelper.createDataFormat().getFormat("$#,##0.00"));

Sheet salarySheet = workbook.createSheet("Salaries");
Row salaryHeaderRow = salarySheet.createRow(0);
Cell salaryNameHeader = salaryHeaderRow.createCell(0);
salaryNameHeader.setCellValue("Name");
Cell salaryPerGoalHeader = salaryHeaderRow.createCell(1);
salaryPerGoalHeader.setCellValue("$ Per Goal");
Cell salarySalaryHeader = salaryHeaderRow.createCell(2);
salarySalaryHeader.setCellValue("AAV");
Cell postsSalaryHeader = salaryHeaderRow.createCell(3);
postsSalaryHeader.setCellValue("Goals + Posts + Crossbars Hit");

int salRowInd = 1;
URL initSalURL = new URL("https://www.capfriendly.com/ajax/cost-per-point/2023/season/all/all/all/costpergoals?p=1");
URLConnection salInitConnection = initSalURL.openConnection();
BufferedReader salIn = new BufferedReader(new InputStreamReader(salInitConnection.getInputStream()));

StringBuilder salContent = new StringBuilder();
String salLine;
while ((salLine = salIn.readLine()) != null) {
  salContent.append(salLine);
  salContent.append(System.lineSeparator());
}

JSONObject initSalJSON = new JSONObject(salContent.toString());
JSONObject initSalJSONData = initSalJSON.getJSONObject("data");
int initSalLimit = initSalJSONData.getInt("count");
salIn.close();


int y = 1;
int salPgCount = initSalLimit;
int previousPerGoal = -1; // Initialize with a large value

  while (salPgCount == 50) {
    URL salURL = new URL("https://www.capfriendly.com/ajax/cost-per-point/2023/season/all/all/all/costpergoals?p=" + y);
    URLConnection salConnection = salURL.openConnection();
    BufferedReader salIn1 = new BufferedReader(new InputStreamReader(salConnection.getInputStream()));

    StringBuilder salContent1 = new StringBuilder();
    String salLine1;
    while ((salLine1 = salIn1.readLine()) != null) {
      salContent1.append(salLine1);
      salContent1.append(System.lineSeparator());
    }
    JSONObject salJSON = new JSONObject(salContent1.toString());
    JSONObject salJSONData = salJSON.getJSONObject("data");

    String html = salJSONData.getString("html");
    

    String startHtmlString = "<html><head><title></title></head><body><table>";
    String endHtml = "</table></body></html>";
    String salHTML = startHtmlString + StringEscapeUtils.unescapeJson(html) + endHtml;
    Document salDoc = Jsoup.parse(salHTML);
    Elements elements = salDoc.body().children();

    Elements trList = elements.get(0).getElementsByTag("tr");
    for (int n = 0; n < trList.size(); n++) {
            String tdValue = trList.get(n).children().get(15).children().text();
            String tdAAVValue = trList.get(n).children().get(7).children().text();
            if(tdValue.isEmpty()){
              salPgCount = 0;
              n = trList.size();
            }
            else{
              String baseName = trList.get(n).children().get(1).children().text();
              String playerName = baseName.substring(baseName.indexOf(" ") + 1) + " " + baseName.substring(0, baseName.indexOf(","));
              int perGoal = Integer.parseInt(tdValue.replace("$", "").replace(",", ""));
              int capHit = Integer.parseInt(tdAAVValue.replace("$", "").replace(",", ""));
              salPgCount = salJSONData.getInt("count");

              Row row2 = salarySheet.createRow(salRowInd);
              Cell nameCell = row2.createCell(0);
              nameCell.setCellValue(playerName);
              Cell perGoalCell = row2.createCell(1);
              perGoalCell.setCellStyle(dollarStyle);
              salarySheet.autoSizeColumn(1);
              perGoalCell.setCellValue(perGoal);
              Cell salCell = row2.createCell(2);
              salCell.setCellValue(capHit);
              salCell.setCellStyle(dollarStyle);
              salRowInd++;
            }
        }
        
        salIn1.close();
        y++;

        
  }
  System.out.println("Salary done");

//#endregion
//^^^^^^^^^^^^^^^^^^^^^  END SALARY SCRIPT ^^^^^^^^^^^^^^^^^^^^^

//vvvvvvSTART ANALYSIS SCRIPTvvvvvvvv
//#region
FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
FormulaEvaluator evaluator1 = workbook.getCreationHelper().createFormulaEvaluator();
for(int i = 0; i<Math.max(limit+1, goalLimit+1); i++)
{
  Row allStatsRow = allStatsSheet.createRow(i);
}
transferSheet(postsSheet, allStatsSheet, 0);
transferSheet(goalSheet, allStatsSheet, 2);
transferSheet(salarySheet, allStatsSheet, 7);



for(int i = 1; i<allStatsSheet.getLastRowNum(); i++)//loops through all of the rows and uses the vlookup function to get the corresponding posts hit stat
{                                                   //then calculates the percentage of goals+posts out of shots taken, formatting it to a percent
  Row allStatsRowRef = allStatsSheet.getRow(i);
  int corrRow = allStatsRowRef.getRowNum()+1;
  Cell vCell = allStatsRowRef.createCell(5, CellType.FORMULA);
  vCell.setCellFormula("VLOOKUP(C"+corrRow+", A:B, 2, FALSE)");
  evaluator.evaluateInCell(vCell);
  Cell salVCell = allStatsRowRef.createCell(10, CellType.FORMULA);
  salVCell.setCellFormula("VLOOKUP(H"+corrRow+", C:F, 4, FALSE) + VLOOKUP(H"+corrRow+", C:F, 3, FALSE)");

  Cell salPerCell = allStatsRowRef.createCell((11), CellType.FORMULA);
  salPerCell.setCellFormula("J"+corrRow + "/K"+corrRow);
  salPerCell.setCellStyle(dollarStyle);
  allStatsSheet.autoSizeColumn(11);
  //evaluator1.evaluateInCell(salVCell);
  
  
  Cell perCell = allStatsRowRef.createCell(6, CellType.FORMULA);
  if(vCell.getCellType()==CellType.NUMERIC){
    perCell.setCellFormula("(E"+corrRow+" + " + vCell.getNumericCellValue() + ")/D"+corrRow);
    evaluator.evaluateInCell(perCell);
    CellStyle stylePercentage = workbook.createCellStyle();
    stylePercentage.setDataFormat(workbook.createDataFormat().getFormat(BuiltinFormats.getBuiltinFormat( 10 )));
    perCell.setCellStyle(stylePercentage);
  }
  else{
    perCell.setCellValue("N/A");
  }
}

//#endregion
//^^^^^^ END ANALYSIS SCRIPT ^^^^^^^^^    
FileOutputStream FileOut = new FileOutputStream("Trocheck Stat.xlsx");
     workbook.write(FileOut);
     System.out.println("File created");
     FileOut.close();
     workbook.close();
     
  }

  //copies data from source sheet to destination sheet starting at a given column
  public static void transferSheet(Sheet sourceSheet, Sheet destinationSheet, int startingColumn)
  {
    Row destHeader = destinationSheet.getRow(0);
    Cell allStatsCumHeader = destHeader.createCell(5);
    allStatsCumHeader.setCellValue("Posts + Crossbars Hit");
    Cell allStatsPercHeader = destHeader.createCell(6);
    allStatsPercHeader.setCellValue("(Goals+Posts)/Shots %");
    Cell allStatsPerPostsHeader = destHeader.createCell(11);
    allStatsPerPostsHeader.setCellValue("$ Per Goals + Posts");
    int rowCount = sourceSheet.getLastRowNum();
    for (int i = 0; i <= rowCount; i++) {
      Row sourceRow = sourceSheet.getRow(i);
      Row destinationRow = destinationSheet.getRow(i);

      int columnCount = sourceRow.getLastCellNum();
      for (int j = 0; j < columnCount; j++) {
          Cell sourceCell = sourceRow.getCell(j);
          Cell destinationCell = destinationRow.createCell(j + startingColumn);

          if (sourceCell.getCellType() == CellType.STRING) {
              destinationCell.setCellValue(sourceCell.getStringCellValue());
              destinationCell.setCellStyle(sourceCell.getCellStyle());
          } 
          else if (sourceCell.getCellType() == CellType.NUMERIC) {
              destinationCell.setCellValue(sourceCell.getNumericCellValue());
              destinationCell.setCellStyle(sourceCell.getCellStyle());
              destinationSheet.autoSizeColumn(j+startingColumn);
          }
      }
    }


  }
}
