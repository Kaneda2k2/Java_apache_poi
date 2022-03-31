# Java_apache_poi
Usos frecuentes apache poi


//create blank workbook
XSSFWorkbook workbook = new XSSFWorkbook();


//Create a blank sheet
 XSSFSheet sheet = workbook.createSheet("Country");
 
// TEST ARRAY
ArrayList<Object[]> data=new ArrayList<Object[]>();
  data.add(new String[]{"Country","Capital","Population"});
  data.add(new Object[]{"India","Delhi",10000});
  data.add(new Object[]{"France","Paris",40000});
  data.add(new Object[]{"Germany","Berlin",20000});
  data.add(new Object[]{"England","London",30000});
  
  
 //Iterate over data and write to sheet
  int rownum = 0;
  for (Object[] countries : data)
  {
   Row row = sheet.createRow(rownum++);
 
   int cellnum = 0;
   for (Object obj : countries)
   {
    Cell cell = row.createCell(cellnum++);
    if(obj instanceof String)
     cell.setCellValue((String)obj);
    else if(obj instanceof Double)
     cell.setCellValue((Double)obj);
    else if(obj instanceof Integer)
     cell.setCellValue((Integer)obj);
    // Setting style only for header
    if(rownum==1)
    {
     CellStyle style=null;
     // Creating a font
        XSSFFont font= workbook.createFont();
        font.setFontHeightInPoints((short)10);
        font.setFontName("Arial");
        font.setColor(IndexedColors.WHITE.getIndex());
        font.setBold(true);
        font.setItalic(false);
 
        style=workbook.createCellStyle();;
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style.setAlignment(CellStyle.ALIGN_CENTER);
        // Setting font to style
        style.setFont(font);
 
        // Setting cell style
        cell.setCellStyle(style);
    }
   }
 
  }
  Row rowGap = sheet.createRow(rownum++);
  Row row = sheet.createRow(rownum++);
  Cell cellTotal = row.createCell(0);
  cellTotal.setCellValue("Total Population");
 
  // Setting cell formula and cell type
  Cell cell = row.createCell(2);
  cell.setCellFormula("SUM(C2:C5)");
  cell.setCellType(Cell.CELL_TYPE_FORMULA);
  try
  {
   //Write the workbook to the file system
   FileOutputStream out = new FileOutputStream(new File("CountriesDetails.xlsx"));
   workbook.write(out);
   out.close();
   System.out.println("CountriesDetails.xlsx has been created successfully");
  }
  catch (Exception e)
  {
   e.printStackTrace();
  }
  finally {
   workbook.close();
  }
 }
}
 
