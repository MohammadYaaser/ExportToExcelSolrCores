import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.*;
import java.text.*;
import java.lang.*;
import java.util.*;

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


class PNExport{

	private static String FILE_NAME = "/webvol/IndexProgram/ExcelExport/MyFirstExcel.xlsx";
		
    public static void main(String args[]){  
		int countParts = 1;
		/* Read Part Numbers from parts.txt file in same directory */
		System.out.println("Stsrated adding data into file\n\n\n\n");
		List<ArrayList<String>> partsLists = new ArrayList<ArrayList<String>>();
		try{
			BufferedReader partsRead = new BufferedReader(new FileReader("parts.txt")); 
			String line;
			ArrayList<String> partsDivided = new ArrayList<String>();
			while ((line = partsRead.readLine()) != null) {
				partsDivided.add(line);
				if(countParts % 250 == 0){
					partsLists.add(partsDivided);
					partsDivided = new ArrayList<String>();
				}
				countParts++;
			}
			partsLists.add(partsDivided);
		}
		catch(Exception exception){
			System.out.println("Exception : " + exception);
		}
		int exportNumber = 1;
		for(ArrayList<String> strArr : partsLists){
			PNExport.export(strArr, exportNumber);
			exportNumber++;
		}
		System.out.println("Completed Succesfully");
	}
	/* Start exporting data  */  
	public static void export(ArrayList<String> partsDivided, int exportNumber){
		String partsQuery ="", solrServer = "stellar.honeywell.com:8989/solr/" , shardQuery = "";

		ArrayList<String> shards = new ArrayList<String>();
		shards.add("ACS-AP_ERP_ORACLE");shards.add("ACS_AML_1S4E");shards.add("ACS_ECAD_MENTOR");shards.add("ACS_MCAD_PDMLink");shards.add("ACS_SAP");shards.add("ECC_AML_ENOVIA");shards.add("ECC_ERP_ORACLE");shards.add("HSG_AML_AGILE");shards.add("HSG_MATERIAL_AGILE");shards.add("HSM_AML_WorkManager");shards.add("HSM_Materials_WorkManager");shards.add("ECRO_ENOVIA");shards.add("HSG_ECRO_AGILE");
		
		Map<String, String> newCoreNameList = new HashMap<String, String>();
		newCoreNameList.put("ACS_AML_1S4E","1S4E AML");
		newCoreNameList.put("ACS-AP_ERP_ORACLE","ORACLE ERP - APAC");
		newCoreNameList.put("HSG_AML_AGILE","Agile AML");
		newCoreNameList.put("ACS_SAP","SAP");
		newCoreNameList.put("HSG_MATERIAL_AGILE","Agile Material");
		newCoreNameList.put("HSG_ECRO_AGILE","Agile ECR/ECO");
		newCoreNameList.put("ACCOLADE","Accolade");
		newCoreNameList.put("ACS_ECAD_MENTOR","Mentor Graphics");
		newCoreNameList.put("ACS_MCAD_PDMLink","PDMLink");
		newCoreNameList.put("ECC_AML_ENOVIA","Enovia AML");
		newCoreNameList.put("ECC_DEVIATIONS_ENOVIA","Enovia Deviation");
		newCoreNameList.put("ECC_Documents_Enovia","Enovia Documents");
		newCoreNameList.put("ECC_ERP_ORACLE","ERP Oracle");
		newCoreNameList.put("ECC_OS_ENOVIA","Enovia OS");
		newCoreNameList.put("ECC_STD_ENOVIA","Enovia STD");
		newCoreNameList.put("ECRO_ENOVIA","Enovia ECRO");
		newCoreNameList.put("ECRO_ROUTES_ENOVIA","Enovia Routes");
		newCoreNameList.put("FOLDER_SEARCH","N/W Folder Search");
		newCoreNameList.put("HSM_AML_WorkManager","WorkManager AML");
		newCoreNameList.put("HSM_Materials_WorkManager","WorkManager Material");

		HashMap<String, ArrayList<String>> coreFieldsAll = new HashMap<String, ArrayList<String>>();
		coreFieldsAll.put("ACCOLADE",new ArrayList<String>(Arrays.asList("VPM_Project_Number", "Accolade_Id", "Project_Name", "Project_Owner_Name", "Description", "Project_Start_Date", "Project_End_Date", "SBG", "GBE", "LOB", "Project_Category", "Project_Classification", "Project_Status", "Project_Type", "Program")));;
		
		coreFieldsAll.put("ACS-AP_ERP_ORACLE",new ArrayList<String>(Arrays.asList("Honeywell_Material_Number", "Manufacturer", "Location(Org)", "Commodity_Manager", "PastUsage", "MonthsUsage", "FutureDemand", "StandardCost", "MaterialCost", "LastRecivedPrice", "OnHand", "DaysSupply", "Description", "Originated", "NewItemCategory", "CategoryOfHBCLOB", "LeadTime", "FixedLotMult", "MinOrderQty", "InventoryPlanningMethod", "MinQuantity", "MaxQuantity", "CQSupplier", "CQPrice", "CQCurrency", "CQCreationDate", "EAVSpendDollars", "Manufacturer", "Buyer", "ProdFamCode", "AnnualUsage", "PortalFlag", "EDIFlag", "InventoryItemStatus", "PriceIndex", "SafetyStockPercent", "FixedDaysSupply", "ConsignedFlag", "ConsignedFromSupplierFlag", "Rev", "UofM", "LastRcvdQty", "LastRcvdDate", "Currancy", "PlannerCode", "CommodityCode", "HBC", "Kanban", "IsHBCLoc", "AssignmentGroup", "ABCClass", "ABCClassDesc", "SBU", "Plant", "Project_Number", "Modified_By")));
		
		coreFieldsAll.put("ACS_AML_1S4E",new ArrayList<String>(Arrays.asList("Honeywell_Material_Number","Description","Manufacturer_Part_Number","Manufacturer","SBU","Additional_Feature","Bus_Driver/Transceiver_Type","Capacitance","Capacitor_Type","Circuit_DC_Voltage-Max","Circuit_RMS_Voltage-Max","Classification","Design_Authority","Energy_Absorbing_Capacity-Max","Family","HPN_Description","ID","IHS_Description","Input_Conditioning","Internal_Items","JESD-30_Code","JESD-609_Code","JESD-95_Code","LC_Status","Lead_Free","Length","Load_Capacitance_(CL)","Manufacturer_Series","Mfr_Pkg_Outline_Code","Moisture_Sensitivity_Level","Mounting_Feature","Multilayer","Name","Negative_Tolerance","Number_of_Bits","Number_of_Functions","Number_of_Ports","Number_of_Potential_Terminals","Number_of_Terminals","Old_Package_Style","Operating_Temperature-Max","Operating_Temperature-Min","Output_Characteristics","Output_Polarity","Package_Body_Material","Package_Code","Package_Shape","Package_Style","Packing_Method","Pardon","Parent_Class_Level_2","Parent_Class_Level_3","Part_Information_Source","Part_Number_Type","Part_Status","Part_Status_(per_Document)","Peak_Reflow_Temperature_(Cel)","Positive_Tolerance","Propagation_Delay_(tpd)","Property_Set_Type","Qualification_Status","Rated_DC_Voltage_(URdc)","Rated_Power_Dissipation_(P)","Rated_Temperature","Reference_Standard","Resistance","Resistor_Type","Seated_Height-Max","Size_Code","Status","Supply_Voltage-Max_(Vsup)","Supply_Voltage-Min_(Vsup)","Supply_Voltage-Nom_(Vsup)","Surface_Mount","Technology","Temperature_Characteristics_Code","Temperature_Coefficient","Temperature_Grade","Terminal_Count_Sequence","Terminal_Finish","Terminal_Form","Terminal_Pitch","Terminal_Placement","Terminal_Position","Terminal_Shape","Time@Peak_Reflow_Temperature-Max_(s)","Timing_Supply_Voltage-Max","Timing_Supply_Voltage-Min","Timing_Temperature-Max","Timing_Temperature-Min","Tolerance","Tri-state_Control_Input","Trigger_Type","Varistor_Type","Vault","Width","Working_Voltage","Years_to_Obsolescence","fmax-Min")));
		
		coreFieldsAll.put("ACS_ECAD_MENTOR",new ArrayList<String>(Arrays.asList("Honeywell_Material_Number", "Manufacturer_Part_Number", "Description", "SBU", "Project_Number", "Modified_By", "Plant")));
		
		coreFieldsAll.put("ACS_MCAD_PDMLink",new ArrayList<String>(Arrays.asList("Honeywell_Material_Number", "Name", "Revision", "Lifecycle_State", "CAD_Type", "Description", "Modified_By", "ECO_Number", "Project_Number", "Material", "Modified_Date", "SBU", "Location", "Url", "Document_Number", "Plant")));

		coreFieldsAll.put("ACS_SAP",new ArrayList<String>(Arrays.asList("Honeywell_Material_Number", "Description", "Honeywell_Material_Type", "Base_Unit_of_Measure", "Lab_Office", "Design_Authority", "LOB", "SBU", "Plant_Material_Status", "Valid_From", "Material_Group", "Gross_Weight", "Net_Weight", "Volume", "Weight_Unit", "Project_Number", "REACH_Relevant", "Material_Category", "Created_By", "Created_On")));

		coreFieldsAll.put("ACS_SAP_PLANT",new ArrayList<String>(Arrays.asList("Honeywell_Material_Number", "Plant", "Procurement_Type", "Special_Procurement", "Plantspmatl_Status", "Valid_From", "MRP_Controller", "Purchasing_Group", "MRP_Type", "SchedMargin_Key", "Planning_Time_Fence", "GR_Processing_Time", "Totrepl_Leadtime", "Planned_Deliv_Time", "Profit_Center", "Country_Of_Origin", "Automatic_PO", "RoHS_Compliance", "RoHS_Compliance_Date", "SIOP_Class", "Service_Level_Days", "Service_Level_Qty", "Fixed_LotSize", "Maximum_LotSize", "Minimum_LotSize", "Rounding_Value", "Storageloc_For_EP", "Prodstor_Location", "Description")));

		coreFieldsAll.put("ECC_AML_ENOVIA",new ArrayList<String>(Arrays.asList("Manufacturer_Part_Number", "Manufacturer", "Material_Alias", "Description", "Manufacturer_Part_Type", "Manufacturer_Part_RoHS_Compliance", "Manufacturer_Part_Peak_Body_Temp", "Manufacturer_Part_PARDON", "Manufacturer_Part_State", "Manufacturer_Part_Revision", "Honeywell_Material_Number", "Design_Authority", "SBU", "Honeywell_Material_Custom_Part", "Honeywell_Material_RoHS_Compliance", "Honeywell_Material_PARDON", "Honeywell_Material_Revision", "Honeywell_Material_State", "comments", "Project_Number", "Modified_By", "Plant")));

		coreFieldsAll.put("ECC_DEVIATIONS_ENOVIA",new ArrayList<String>(Arrays.asList("Description", "Analysis", "Permanent_Corrective_Action", "Honeywell_Material_Number", "Document_Number", "Deviation_Name", "Part_Spec_Number_to_Deviate", "Part_Spec_Number_to_Deviate_to", "Product", "create_Date", "Deviation_Expiration", "Estimated_Implementation_Date", "Originator", "Responsible_for_Corrective_Action", "Planner", "Modified_By", "Deviation_Type", "Deviated_Quantity", "Build_Department", "Monthly_Usage", "Conditions", "Current_Inventory_Units", "State", "SBU", "Reason", "Product_Line", "ECR_Evaluator", "Design_Authority", "ECO_Facilitator", "Project_Number", "Design_Engineer", "Marketing_Name", "LOB", "Region", "Lab_Office", "Material_Group")));

		coreFieldsAll.put("ECC_Documents_Enovia",new ArrayList<String>(Arrays.asList("SBU", "Project_Number", "Document_Number", "Material_Alias", "Modified_By", "Originator", "Description", "Document_Type", "Design_Authority", "Model", "Legacy_DocType", "Legacy_DocRevision", "Keywords", "Keyword", "Source", "Resolution", "External_Issue", "First_UsedOn", "LOB", "OEM", "Comment", "Document_Revision", "Document_State", "DQA_Engineer", "InWork_Date_Actual", "Active_Date_Actual")));

		coreFieldsAll.put("ECC_ERP_ORACLE",new ArrayList<String>(Arrays.asList("Description", "Honeywell_Material_Number", "Location(Org)", "CommodityManager", "PastUsage", "MonthsUsage", "FutureDemand", "StandardCost", "MaterialCost", "LastRecivedPrice", "OnHand", "DaysSupply", "CreationDate", "NewItemCategory", "CategoryOfHBCLOB", "LeadTime", "FixedLotMult", "MinOrderQty", "InventoryPlanningMethod", "MinQuantity", "MaxQuantity", "CQSupplier", "CQPrice", "CQCurrency", "CQCreationDate", "EAVSpendDollars", "Buyer", "ProdFamCode", "AnnualUsage", "PortalFlag", "EDIFlag", "InventoryItemStatus", "PriceIndex", "SafetyStockPercent", "FixedDaysSupply", "ConsignedFlag", "ConsignedFromSupplierFlag", "Rev", "UofM", "LastRcvdQty", "LastRcvdDate", "Currancy", "PlannerCode", "CommodityCode", "HBC", "Kanban", "IsHBCLoc", "AssignmentGroup", "ABCClass", "ABCClassDesc", "SBU", "Project_Number", "Modified_By", "Plant")));

		coreFieldsAll.put("ECC_ERP_SAP",new ArrayList<String>(Arrays.asList("Honeywell_Material_Number", "Description", "Lab/Office", "Document_Number", "Document_Version", "Plant_Material_Status", "Plant", "Honeywell_Material_RoHS_Compliance", "Minimum_Lot_Size", "Price_Unit", "Standard_Price", "currency", "Vendor_Number", "EAN/UPC", "PIR_Number", "Plant_Delivery_Time", "Purchasing_Group", "Net_Price", "Standard_Quantity", "Minimum_Quantity", "SBU", "Project_Number", "Modified_By", "Body_Breadth", "Body_Length_or_Diameter", "Body_Height", "Classification", "First_Element_Resistance", "JESD-609_Code", "Manufacturer_Series", "Mfr_Package_Outline_Code", "Mounting_Feature", "Operating_Temperature-Max", "Operating_Temperature-Min", "Package_Code", "Package_Shape", "Package_Style", "Old_Package_Style", "Physical_Dimension", "Shape/Size_Description", "Size", "Size_Code", "Vault", "Wire_Gauge", "Design_Authority", "Manufacturer", "Honeywell_Material_PARDON", "Manufacturer_Part_PARDON", "Manufacturer_Part_Type", "Manufacturer_Part_RoHS_Compliance", "Manufacturer_Part_Peak_Body_Temp", "Honeywell_Material_Custom_Part", "Parent_Class_Level_2", "Parent_Class_Level_3", "Pardon", "Terminal_Pitch", "Rated_DC_Voltage_(URdc)", "Capacitance", "Capacitor_Type", "Rated_Power_Dissipation_(P)", "Working_Voltage")));

		coreFieldsAll.put("ECC_OS_ENOVIA",new ArrayList<String>(Arrays.asList("Honeywell_Material_Number", "SBU", "Project_Number", "Document_Number", "Material_Alias", "Honeywell_Material_Revision", "Honeywell_Material_State", "Modified_By", "Custom_Part", "Comment", "EAN", "LOB", "Design_Authority", "Honeywell_Material_RoHSCompliance", "Product_Customer_Name", "Product_Customer_Part_Number", "Product_Family_and_Model_Code", "Product_Model_Description", "Production_Make_Buy_Code", "Universal_Product_Code", "Pack_Code", "Pack_Type", "Pack_Indicator", "Description")));

		coreFieldsAll.put("ECC_STD_ENOVIA",new ArrayList<String>(Arrays.asList("Document_Number", "Document_Type", "Document_Revision", "Description", "related_Manufacturer_Part_Number", "related_Manufacturer_Part_Type", "related_Manufacturer_Part_Description", "related_Honeywell_Material_Number", "related_Honeywell_Material_Type", "Design_Authority", "SBU", "Source", "comments", "Project_Number", "Modified_By", "Plant")));

		coreFieldsAll.put("ECRO_ENOVIA",new ArrayList<String>(Arrays.asList("SBU", "ECRO_Name", "State", "Project_Number", "ECO_Facilitator", "ECR_Evaluator", "Document_Number", "Honeywell_Material_Number", "Modified_By", "Description", "Design_Authority", "Design_Engineer", "Originator", "Product_Line", "Marketing_Name", "LOB", "Region", "related_ECRO", "create_Date", "modified_Date", "create_Date_Actual", "submit_Date_Actual", "evaluate_Date_Actual", "review_Date_Actual", "planECO_Date_Actual", "complete_Date_Actual", "defineComp_Date_Actual", "designWork_Date_Actual", "release_Date_Actual", "rejected_Date_Actual", "implemented_Date_Actual", "cancelled_Date_Actual", "promote_CountToCreate", "promote_CountToSubmit", "promote_CountToEvaluate", "promote_CountToReview", "promote_CountToPalnEco", "promote_CountToComplete", "demote_CountToCreate", "demote_CountToSubmit", "demote_CountToEvaluate", "demote_CountToReview", "demote_CountToPalnEco", "demote_CountToComplete", "promote_CountToDefineComp", "promote_CountToDesignWork", "promote_CountToRelease", "promote_CountToImplemented", "demote_CountToDefineComp", "demote_CountToDesignWork", "demote_CountToRelease", "demote_CountToImplemented", "demote_CountTotal", "promote_CountTotal")));

		coreFieldsAll.put("ECRO_ROUTES_ENOVIA",new ArrayList<String>(Arrays.asList("ECRO_Name", "Route_Name", "Route_State", "Route_Owner", "Task_Name", "Task_Instructions", "Task_Owner", "Task_Assigne_Comments", "Task_State", "Task_DueDate", "Task_DueDate_Offset", "Route_Content", "Design_Authority", "Description")));

		coreFieldsAll.put("FOLDER_SEARCH",new ArrayList<String>(Arrays.asList("title", "size", "location", "file_type", "modified_date", "server_name")));

		coreFieldsAll.put("HSG_AML_AGILE",new ArrayList<String>(Arrays.asList("Honeywell_Material_Number", "Manufacturer", "Manufacturer_Part_Number", "Manufacturer_Part_Type", "Manufacturer", "Manufacturer_Part_RoHS_Compliance", "Manufacturer_Part_State", "Manufacturer_Part_Revision", "Material_Alias", "1SFE", "SHVC_List", "Weight_mg", "SHVC_Presence", "PPL_Compliant", "EndOfLifeDate", "YearsToObsolete", "Modified_By", "Originated", "SBU", "Description", "Project_Number", "Plant")));

		coreFieldsAll.put("HSG_ECRO_AGILE",new ArrayList<String>(Arrays.asList("Description", "Honeywell_Material_Number", "ECRO_Name", "create_Date", "Originator", "Modified_By", "Change_Type", "Work_Flow", "State", "SBU", "Reason", "ECO_Facilitator", "ECR_Evaluator", "Project_Number", "Design_Authority", "Design_Engineer", "Product_Line", "Marketing_Name", "LOB", "Region")));

		coreFieldsAll.put("HSG_MATERIAL_AGILE",new ArrayList<String>(Arrays.asList("BusinessLine", "Honeywell_Material_Type", "Honeywell_Material_Number", "Description", "Honeywell_Material_State", "Honeywell_Material_Revision", "ODM_Code", "Document_Number", "AgencyApproval", "DateCreated", "CAD_SymbolStatus", "Modified_By", "Originated", "Design_Authority", "EPL_Owner", "RoHS_DeclarationByDesignCentre", "RoHS_RollupResult", "RoHS_RollupDate", "SafetyCritical", "UOM", "PartCategory", "RevIncorpDate", "RevReleaseDate", "Custom_COTS", "SBU", "MRP_Type", "Manufacturer_Part_Type", "Project_Number", "Plant")));

		coreFieldsAll.put("HSM_AML_WorkManager",new ArrayList<String>(Arrays.asList("Honeywell_Material_Number", "Manufacturer_Part_Number", "Manufacturer", "Manufacturer_Part_Type", "Manufacturer_Part_RoHS_Compliance", "Manufacturer_Part_State", "Manufacturer_Part_Revision", "Material_Alias", "1SFE", "SHVC_List", "Weight_mg", "SHVC_Presence", "PPL_Compliant", "EndOfLifeDate", "YearsToObsolete", "Modified_By", "Originated", "SBU", "Description", "Project_Number", "Plant")));

		coreFieldsAll.put("HSM_Materials_WorkManager",new ArrayList<String>(Arrays.asList("Honeywell_Material_Type", "Honeywell_Material_Number", "Description", "Honeywell_Material_Revision", "Modified_By", "SBU")));


		for (String line: partsDivided) {
			partsQuery += "'"+ line + "'" + " OR "; 
		}
		partsQuery = partsQuery.substring(0, partsQuery.length() - 4);
		/* Start feching data from solr */
		SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yyyy HH:mm:ss");  
    	Date date = new Date();  
		FILE_NAME = "/webvol/IndexProgram/ExcelExport/MyFirstExcel_" + exportNumber + ".xlsx";
		//FILE_NAME = "/webvol/IndexProgram/ExcelExport/MyFirstExcel.xlsx";
    	HttpSolrServer server =new HttpSolrServer("http://" + solrServer + "UniCore");

        SolrQuery query = new SolrQuery();
        query.setQuery("(Honeywell_Material_Number:(" + partsQuery + "))");
        
        query.setStart(0);
		query.set("defType", "edismax");
		for(String sh : shards) 
		{
			if(!shardQuery.equals("")) 
			{
				shardQuery += ","; 
			}
			shardQuery += solrServer + sh;
		}
		//shardQuery = solrServer +"ACS_SAP";
		System.out.println(shardQuery);
		query.set("shards", shardQuery);
		//query.set("fl", "Honeywell_Material_Number,Lab_Office");
		//query.setRows(10);
		query.setRows(99999999);

		System.out.println("Started Exporting : " + FILE_NAME + "Adding :" + partsDivided.size());
		XSSFWorkbook workbook=null;
		XSSFSheet sheet;
		FileOutputStream out = null;
		int rowCount = 0;
        // SXSSFWorkbook wb = null;
		// FileOutputStream fos = null;
		// int rowNum = 0;
		try {
			/* XSSF  */
			FileInputStream file = new FileInputStream(new File("/webvol/IndexProgram/ExcelExport/MyFirstExcel.xlsx"));

			//Create Workbook instance holding reference to .xlsx file
			workbook = new XSSFWorkbook(file);

			//Get first/desired sheet from the workbook
			//Most of people make mistake by making new sheet by looking in tutorial
			sheet = workbook.getSheetAt(workbook.getActiveSheetIndex());

			//Get the count in sheet
			rowCount = sheet.getLastRowNum();
			System.out.println("Available Rows in sheet :" + rowCount);
		 	/*  */
			// keep 100 rows in memory, exceeding rows will be flushed to disk
			// wb = new SXSSFWorkbook(SXSSFWorkbook.DEFAULT_WINDOW_SIZE/* 100 */);
			// Sheet sh = wb.createSheet();
			// @SuppressWarnings("unchecked")

	        long start = System.currentTimeMillis();
	        //System.out.println("query : " + query);
			QueryResponse response = server.query(query);

	        SolrDocumentList results = response.getResults();
			Iterator<SolrDocument> solrDocumentIterator = results.iterator();
			

			// adding 
			HashMap<String, ArrayList<ArrayList<String>>> multiMap = new HashMap<String, ArrayList<ArrayList<String>>>();
		
			ArrayList<ArrayList<String>> eachCoreData = new ArrayList<ArrayList<String>>();

			while(solrDocumentIterator.hasNext()) {
				ArrayList<String> materialDataRow = new ArrayList<String>();
				SolrDocument doc = solrDocumentIterator.next();
				String shard = doc.getFieldValue("[shard]").toString().replace(solrServer,"");
				String shardName = newCoreNameList.get(shard);
				materialDataRow.add(shardName);
				//materialDataRow.add(shard);
				for(String field : coreFieldsAll.get(shard)) {
					String cellValue = "";
					if(doc.getFieldValue(field) != null)
					{
						cellValue = doc.getFieldValue(field).toString();
					}
					materialDataRow.add(cellValue);
				}
				eachCoreData.add(materialDataRow);
			}

			for(String sh : shards) {
				ArrayList<ArrayList<String>> materialsDataPush = new ArrayList<ArrayList<String>>();
				for (int i = 0; i < eachCoreData.size(); i++) {

					//String coreName = (eachCoreData.get(i).get(0)).toString().replaceAll(" ","_");
					String coreName = (eachCoreData.get(i).get(0)).toString();
					if((newCoreNameList.get(sh) == coreName ) || ( (newCoreNameList.get(sh)).equals(coreName)) )
					{
						materialsDataPush.add(eachCoreData.get(i));
					}			
				}
				if(materialsDataPush.size() > 0 )
				{
					multiMap.put(sh,materialsDataPush);
				}
			}
			int rowNo = 0;
			Row row = null;
			Cell cell = null;
			int colNum = 0;

			for(Map.Entry<String, ArrayList<ArrayList<String>>> entry : multiMap.entrySet()) {

				String coreName = entry.getKey();
				String coreNameDisplay = newCoreNameList.get(coreName).replaceAll("/"," or ");
				//coreName = getCoreName(coreName);
				//Sheet shName = wb.createSheet(coreName);
				if(workbook.getSheet( coreNameDisplay ) != null){
					sheet = workbook.getSheet( coreNameDisplay ) ;
				}
				else{
					sheet = workbook.createSheet( coreNameDisplay );
				}
				//sheet = workbook.getSheetAt(workbook.getActiveSheetIndex());

				//Get the count in sheet
				rowCount = sheet.getLastRowNum();
				System.out.println("Available Rows in sheet :" + rowCount);
					//Sheet shName = wb.createSheet( coreNameDisplay );
				//@SuppressWarnings("unchecked")

				ArrayList<ArrayList<String>> coreData = entry.getValue();
				rowNo = 0;
				row = sheet.createRow(rowCount++);
				colNum = 0;
				row.createCell(colNum++).setCellValue("Data Sources");
				for(int i = 0; i < coreFieldsAll.get(coreName).size(); i++) {
					row.createCell(colNum++).setCellValue((String) coreFieldsAll.get(coreName).get(i).toString().replaceAll("_"," "));
				}

				for (int i = 0; i < coreData.size(); i++) {
					row = sheet.createRow(rowCount++);
					colNum = 0;
					for (int j = 0; j < coreData.get(i).size(); j++) {
						String field = (coreData.get(i).get(j)).toString();
						cell=row.createCell(colNum++);
						if( field.length() < 0)
						{
							cell.setCellValue("NULL/EMPTY");
							//cell.setCellType(Cell.CELL_TYPE_BLANK);
						}
						else{
							cell.setCellValue((String) coreData.get(i).get(j));
						}
					}
				}
			}
		// ends
	        /* System.out.println("solr results " + results.size() + ", " + results.getNumFound() + "\n\n\nmultiMap" + multiMap);
	        SolrDocument doc = null;
	        Collection<String> fields = null;
			Iterator<String> fieldIterator = null;
			int colNum = 0;

			Row row = sheet.createRow(rowCount++);
			//row.createCell(colNum++).setCellValue("Honeywell_Material_Number");
			//row.createCell(colNum++).setCellValue("Lab_Office");

	        while(solrDocumentIterator.hasNext()) {
				doc = solrDocumentIterator.next();
				String hnpValue = (doc.getFieldValue("Honeywell_Material_Number")).toString();
				if( partsDivided.contains(hnpValue) )
				{
					fields = doc.getFieldNames();
					fieldIterator = fields.iterator();
					Row partRow = sheet.createRow(rowCount++);
					colNum = 0;

					while(fieldIterator.hasNext()) {
						String field = (doc.getFieldValue(fieldIterator.next())).toString();
						if( field.length() < 32767){
							partRow.createCell(colNum++).setCellValue((String) field );
						}
						else{
							System.out.println("more than 32767");
						}
					}
				}
			} */
			
			out = new FileOutputStream(new File("/webvol/IndexProgram/ExcelExport/MyFirstExcel.xlsx"));
            workbook.write(out);
            out.close();
            System.out.println("Update Successfully");

		} catch (Exception ex) {
			System.out.println("Exception in :" + ex + "at rowNum :" + rowCount);
		} finally {
			try {
				if (out != null) {
					out.close();
				}
			} catch (IOException e) {
			}
			try {
				if (workbook != null) {
					workbook.close();
				}
			} catch (IOException e) {
			}
		}
		/* Export ens here */
	}
}
