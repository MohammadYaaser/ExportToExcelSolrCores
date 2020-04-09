import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;


import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Sheet;
import java.io.File;

import java.text.SimpleDateFormat;  
import java.util.Date; 
import java.util.ArrayList; 
import java.util.Collection; 
import java.util.Iterator; 


import org.apache.solr.client.solrj.impl.HttpSolrServer;
import org.apache.solr.client.solrj.response.FacetField;
import org.apache.solr.client.solrj.response.FieldStatsInfo;
import org.apache.solr.client.solrj.response.QueryResponse;
import org.apache.solr.client.solrj.SolrQuery;
import org.apache.solr.client.solrj.SolrServer;
import org.apache.solr.common.SolrDocument;
import org.apache.solr.common.SolrDocumentList;


class ExportToExcel{  

    private static final String FILE_NAME = "/webvol/IndexProgram/ExcelExport/MyFirstExcel.xlsx";

    public static void main(String args[]){  
    	SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yyyy HH:mm:ss");  
    	Date date = new Date();  

    	HttpSolrServer server =new HttpSolrServer("http://qastellar.honeywell.com:8983/solr/UniCore");
    	//SolrClient client = new HttpSolrClient.Builder("http://stellar.honeywell.com:8989/solr/UniCoreMaterial").build();

        SolrQuery query = new SolrQuery();
        query.setQuery("text:5");

        query.setStart(0);
        query.set("defType", "edismax");
        //query.set("shards","qastellar.honeywell.com:8983/solr/ACS_SAP,qastellar.honeywell.com:8983/solr/ECC_ERP_ORACLE,qastellar.honeywell.com:8983/solr/HSG_AML_AGILE,qastellar.honeywell.com:8983/solr/FOLDER_SEARCH,qastellar.honeywell.com:8983/solr/ACS-AP_ERP_ORACLE,qastellar.honeywell.com:8983/solr/ACCOLADE,qastellar.honeywell.com:8983/solr/ECC_OS_ENOVIA,qastellar.honeywell.com:8983/solr/HSG_ECRO_AGILE,qastellar.honeywell.com:8983/solr/ACS_AML_1S4E,qastellar.honeywell.com:8983/solr/ECC_AML_ENOVIA,qastellar.honeywell.com:8983/solr/ECC_STD_ENOVIA,qastellar.honeywell.com:8983/solr/HSG_MATERIAL_AGILE,qastellar.honeywell.com:8983/solr/ACS_ECAD_MENTOR,qastellar.honeywell.com:8983/solr/ECC_DEVIATIONS_ENOVIA,qastellar.honeywell.com:8983/solr/ECRO_ENOVIA,qastellar.honeywell.com:8983/solr/HSM_AML_WorkManager,qastellar.honeywell.com:8983/solr/ACS_MCAD_PDMLink,qastellar.honeywell.com:8983/solr/ECC_Documents_Enovia,qastellar.honeywell.com:8983/solr/ECRO_ROUTES_ENOVIA,stellar.honeywell.com:8989/solr/HSM_Materials_WorkManager");
        query.set("shards","qastellar.honeywell.com:8983/solr/ACS_SAP,qastellar.honeywell.com:8983/solr/ECC_ERP_ORACLE,qastellar.honeywell.com:8983/solr/HSG_AML_AGILE,qastellar.honeywell.com:8983/solr/ACS-AP_ERP_ORACLE,qastellar.honeywell.com:8983/solr/ACCOLADE,qastellar.honeywell.com:8983/solr/ACS_AML_1S4E,qastellar.honeywell.com:8983/solr/ECC_AML_ENOVIA,qastellar.honeywell.com:8983/solr/HSG_MATERIAL_AGILE,qastellar.honeywell.com:8983/solr/ACS_ECAD_MENTOR,qastellar.honeywell.com:8983/solr/HSM_AML_WorkManager,qastellar.honeywell.com:8983/solr/ACS_MCAD_PDMLink,qastellar.honeywell.com:8983/solr/HSM_Materials_WorkManager,qastellar.honeywell.com:8983/solr/FOLDER_SEARCH,qastellar.honeywell.com:8983/solr/ECC_OS_ENOVIA,qastellar.honeywell.com:8983/solr/HSG_ECRO_AGILE,qastellar.honeywell.com:8983/solr/ECC_STD_ENOVIA,qastellar.honeywell.com:8983/solr/ECC_DEVIATIONS_ENOVIA,qastellar.honeywell.com:8983/solr/ECRO_ENOVIA,qastellar.honeywell.com:8983/solr/ECC_Documents_Enovia,qastellar.honeywell.com:8983/solr/ECRO_ROUTES_ENOVIA");
        query.setRows(100000);

/*
        try{

        	QueryResponse response = server.query(query);

	        SolrDocumentList results = response.getResults();
	        Iterator<SolrDocument> solrDocumentIterator = results.iterator();
	        System.out.println("solr results " + results.size());

	        while(solrDocumentIterator.hasNext()) {
				SolrDocument doc = solrDocumentIterator.next();
				Collection<String> fields = doc.getFieldNames();
				Iterator<String> fieldIterator = fields.iterator();
				while(fieldIterator.hasNext()) {
					System.out.println(doc.getFieldValue(fieldIterator.next()));
				}
			}
        }
        catch(Exception ex){
        	System.out.println("solr Exception :" + ex );
        }*/
        

     	System.out.println("Hello Java");

        SXSSFWorkbook wb = null;
		FileOutputStream fos = null;
		int rowNum = 0;
		try {
			// keep 100 rows in memory, exceeding rows will be flushed to disk
			wb = new SXSSFWorkbook(SXSSFWorkbook.DEFAULT_WINDOW_SIZE/* 100 */);
			Sheet sh = wb.createSheet();
			@SuppressWarnings("unchecked")

			
	        long start = System.currentTimeMillis();
	        System.out.println("query : " + query);
			QueryResponse response = server.query(query);

	        SolrDocumentList results = response.getResults();
	        Iterator<SolrDocument> solrDocumentIterator = results.iterator();
	        System.out.println("solr results " + results.size() + ", " + results.getNumFound());
	        SolrDocument doc = null;
	        Collection<String> fields = null;
	        Iterator<String> fieldIterator = null;

	        while(solrDocumentIterator.hasNext()) {
				doc = solrDocumentIterator.next();
				fields = doc.getFieldNames();
				fieldIterator = fields.iterator();

				Row row = sh.createRow(rowNum++);
		        int colNum = 0;
				
				while(fieldIterator.hasNext()) {
					//Cell cell = row.createCell(colNum++);
					//cell.setCellValue((String) doc.getFieldValue(fieldIterator.next()));
					//System.out.println(doc.getFieldValue(fieldIterator.next()));
					 String field = (doc.getFieldValue(fieldIterator.next())).toString();
					//System.out.println(field);
					if( field.length() < 32767)
					{
						row.createCell(colNum++).setCellValue((String) field );
					}
		            else{
		            	System.out.println("more than 32767");
		            	row.createCell(colNum++).setCellValue((String) "more than 32767 not allowed" );
		            }
				}
			}

/*			Object[][] datatypes = {
		        {"Head First Java", "Kathy Serria", "0079", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria"},
		        {"Head First Java", "Kathy Serria", "0079", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria"},
		        {"Head First Java", "Kathy Serria", "0079", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria"},
		        {"Head First Java", "Kathy Serria", "0079", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria", "Head First Java", "Kathy Serria"},
		    };

	        
	        System.out.println("Creating excel at : " + start);
		    //for(int i = 0; i < 40000; i++) 
		    {   
		        for (Object[] datatype : datatypes) {
		            Row row = sh.createRow(rowNum++);
		            int colNum = 0;
		            for (Object field : datatype) {
		                Cell cell = row.createCell(colNum++);
		                cell.setCellValue((String) field);
		                // if (field instanceof String) {
		                //     cell.setCellValue((String) field);
		                // } else if (field instanceof Integer) {
		                //     cell.setCellValue((Integer) field);
		                // }
		            }
		        }
		    }*/
			fos = new FileOutputStream(FILE_NAME);
			wb.write(fos);
			long end = System.currentTimeMillis();
			float sec = (end - start) / 1000F; System.out.println(sec + " seconds");
	        System.out.println("Completed excel at : " + sec);

		} catch (Exception ex) {
			System.out.println("Exception in :" + ex + "at rowNum :" + rowNum);
		} finally {
			try {
				if (fos != null) {
					fos.close();
				}
			} catch (IOException e) {
			}
			try {
				if (wb != null) {
					wb.close();
				}
			} catch (IOException e) {
			}
		}

    }  
}
