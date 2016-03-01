package genericclasses;



import configuration.Resourse_path;

import java.io.BufferedInputStream;
import java.io.DataInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Hashtable;
import java.util.Map;
import java.util.TreeMap;

import configuration.Resourse_path;
import driver.Driver;

	public class BenefitAsiaFunction {
		
		private static final String NewReCount = null;

		Hashtable<String, String> HRdata = new Hashtable<String, String>();
		Map<String, String> HRmap = new TreeMap<String, String>();
		
//		public static String empcsv = Resourse_path.csv_path+"/AutomationSG/Employee_1.csv";;
//		public static String depcsv = Resourse_path.csv_path+"/AutomationSG/Dependent_1.csv";
		
		public String new_empcsv = Resourse_path.csv_path+"/AutomationSG/output_" + Driver.Empcsvfname;
		public String new_depcsv = Resourse_path.csv_path+"/AutomationSG/output_"+Driver.Depcsvfname;
		public String new_historycsv = Resourse_path.csv_path+"/AutomationSG/output_"+Driver.Histcsvfname;
		

		public static void main(String[] args) throws IOException, InterruptedException {
			// TODO Auto-generated method stub
//			UpdateEmpDepData cv = new UpdateEmpDepData();
//			cv.generateCode(empcsv,depcsv);
		}	
		
		public void generateCode(String empcsv, String depcsv, String historycsv) throws IOException{
			if (empcsv!=null)
			
			{
				updateEmpdata(empcsv);
			}
			if (depcsv!=null)
			
			{
				updateDep(depcsv);
			}
			if (historycsv!=null)
			{
					updateHistoryData(historycsv);
			}
			deleteFile(empcsv,depcsv, historycsv);
			renamfile(new_empcsv,empcsv,new_depcsv,depcsv, new_historycsv, historycsv);
		}
		
		public void renamfile( String new_empcsv, String empcsv,String new_depcsv, String depcsv, String new_historycsv, String historycsv){
		//rename file
			if (empcsv!="")
			{
				File emp = new File(new_empcsv);
			//rename emp file
				if(emp.exists())
				{
				emp.renameTo(new File(empcsv));
				}
			}
			
			//rename dep file
			if (depcsv!="")
			{
				File Dep = new File(new_depcsv);
				if(Dep.exists())
				{
				Dep.renameTo(new File(depcsv));	
				}
			}
			if (historycsv!=null)
			{
				File historydata = new File(new_historycsv);
				//rename history file
				if(historydata.exists())
				{
					historydata.renameTo(new File(historycsv));
				}
			}
		}
		
		public void deleteFile(String empcsv, String depcsv, String historycsv){
			String [] filename = {empcsv,depcsv,historycsv};
			for(String x:filename){
				if(x != null)
				{
					File file = new File(x);
					if(file.exists())
					{
					file.delete();
					}
				}
			}
		}
		
		@SuppressWarnings("deprecation")
		public void updateEmpdata(String fName) throws IOException{
			String thisLine;
			int rowno = 0;
			FileInputStream fis = new FileInputStream(fName);
			DataInputStream myInput = new DataInputStream(fis);
			
			ArrayList<String> newdata = new ArrayList<String>();
			
			while ((thisLine = myInput.readLine()) != null) {
				ArrayList<String> olddata = new ArrayList<String>();
				
				String strar[] = thisLine.split(",");
				for(int m=0; m<strar.length; m++){
					olddata.add(strar[m]);
				}
				
				if (rowno != 0) {
					String old_HRempID = olddata.get(0);
					String newHRID = getNewHRId(fName, old_HRempID);
					// System.out.println("old "+old_HRempID+" replaced with : "+newHRID);
					olddata.set(0, newHRID);
					olddata.set(1, newHRID);
					olddata.set(2, newHRID);
					olddata.set(7, newHRID);
					HRdata.put(old_HRempID, newHRID);
				}
				
				StringBuilder finaldata = new StringBuilder();
				
				for(int i=0; i<olddata.size(); i++){
					finaldata.append(olddata.get(i)+",");
				}
				
				String after_rem = finaldata.substring(0, finaldata.length()-1);
				newdata.add(after_rem);
				//assigning row
				rowno++;
			}	
			myInput.close();
			HRmap = new TreeMap<String, String>(HRdata); 
//			System.out.println("\n");
//			System.out.println("Employee details ");
//			System.out.println(" \n");
			for(int i=0; i<newdata.size(); i++){
//				System.out.println(newdata.get(i));
			}
			//write data to csv file
			FileWriter writer = new FileWriter(new_empcsv);
			for (int i=0;  i < newdata.size(); i++) {
				writer.append(newdata.get(i));
				writer.append('\n');
			}
		    writer.flush();
		    writer.close();
			
		}

		public void updateDep(String fName) throws IOException{
			ArrayList<String> newdata = new ArrayList<String>();
			int row = 0;		
			
			for(Map.Entry<String,String> map:HRmap.entrySet()){  
				   
				   String oldhr = map.getKey().toString();
				   String newhr = map.getValue().toString();

					String thisLine;
					FileInputStream fis = new FileInputStream(fName);
					DataInputStream myInput = new DataInputStream(fis);
			
					while ((thisLine = myInput.readLine()) != null) {
						
						ArrayList<String> olddata = new ArrayList<String>();
						String strar[] = thisLine.split(",");
						for(int m=0; m<strar.length; m++){
							olddata.add(strar[m]);
						}
						String HRID = olddata.get(0);
						if(row==0){
							StringBuilder finaldata = new StringBuilder();
							for(int i=0; i<olddata.size(); i++){
								finaldata.append(olddata.get(i)+",");
							}
							String after_rem = finaldata.substring(0, finaldata.length()-1);
							newdata.add(after_rem);
						}
						
//						System.out.println(oldhr+ " compared to "+HRID);
						if(oldhr.equals(HRID)){
							olddata.set(0, newhr);
							String FName = olddata.get(2);
							String Fname_Split[] = FName.split("_");
							String FnameNew = newhr + "_" + Fname_Split[3];
							String NIDTemp = olddata.get(6);
							String NID_Split[] = NIDTemp.split("_");
							String NIDNew = newhr+ "_" + NID_Split[3];
							String appendeddata = olddata.get(7);
							olddata.set(2, FnameNew);
							olddata.set(6, NIDNew+"_NID");
							
							StringBuilder finaldata = new StringBuilder();
							for(int i=0; i<olddata.size(); i++){
								finaldata.append(olddata.get(i)+",");
							}
							String after_rem = finaldata.substring(0, finaldata.length()-1);
							newdata.add(after_rem);
						}
						
						//assigning row
						row++;
					}	
					myInput.close();
			 }
//			System.out.println("\n");
//			System.out.println("Dependent details ");
//			System.out.println(" \n");
			
			for (int i=0;  i < newdata.size(); i++) {
//				System.out.println(newdata.get(i));
			}

			//write data to csv file
			FileWriter writer = new FileWriter(new_depcsv);
			for (int i=0;  i < newdata.size(); i++) {
				writer.append(newdata.get(i));
				writer.append('\n');
			}
		    writer.flush();
		    writer.close();
		}

		
		public String getNewHRId(String file, String old_hrid) throws IOException{
			String hrid = "";
			String newString ="";
			
								
			String tempHead = old_hrid.toLowerCase();
        	      		
        	String temp = tempHead;
        		
        	String[] newArrTemp = temp.split("_");
			
        	newString = newArrTemp[0]+"_"+newArrTemp[1]+"_";
        	
			
			//String brk_oldhrid = old_hrid.substring(0, 9);
			String brk_oldhrid = newString;
			String brk_oldhridnum = newArrTemp[2];
			
			int num = Integer.parseInt(brk_oldhridnum);
			int totalrow = count(file);
			
			for(int i=0; i<totalrow; i++){
				String genID = brk_oldhrid+num;
//				System.out.println(" genID "+genID);
				hrid = genID;
				boolean st = returnRdStatus(file, genID);
				if(st==false){
					break;
				}
				num++;
			}
			return hrid;
		}
		
		public boolean returnRdStatus(String file, String old_hrid){
			boolean rt = false;
			String thisLine;
			try{
				ArrayList<String> storageID = new ArrayList<String>();
				storageID.add(old_hrid);
				
				if(storageID.size()>0){
					if(storageID.contains(old_hrid)){
						rt = true;
					}
				}
				
				FileInputStream fis = new FileInputStream(file);
				DataInputStream myInput = new DataInputStream(fis);
				int row = 0;
				while ((thisLine = myInput.readLine()) != null) {
					if(row!=0){
						String strar[] = thisLine.split(",");
						String orghrid = strar[0];
						if(old_hrid.equals(orghrid)){
							rt = true;
							break;
						}	
					}
					row++;
				}	
				myInput.close();
			}catch(Exception e){
				e.printStackTrace();
			}
			return rt;
		}
		
		public int count(String filename) throws IOException {
		    InputStream is = new BufferedInputStream(new FileInputStream(filename));
		    try {
		    byte[] c = new byte[1024];
		    int count = 0;
		    int readChars = 0;
		    boolean empty = true;
		    while ((readChars = is.read(c)) != -1) {
		        empty = false;
		        for (int i = 0; i < readChars; ++i) {
		            if (c[i] == '\n') {
		                ++count;
		            }
		        }
		    }
		    return (count == 0 && !empty) ? 1 : count;
		    } finally {
		    is.close();
		   }
		}

		/*
		
		public void updateHistoryData(String fName) throws IOException{
			ArrayList<String> newdata = new ArrayList<String>();
			int row = 0;		
			
			for(Map.Entry<String,String> map:HRmap.entrySet())
			{  
				   
				   String oldhr = map.getKey().toString();
				   String newhr = map.getValue().toString();

					String thisLine;
					FileInputStream fis = new FileInputStream(fName);
					DataInputStream myInput = new DataInputStream(fis);
			
					while ((thisLine = myInput.readLine()) != null) {
						
						ArrayList<String> olddata = new ArrayList<String>();
						String strar[] = thisLine.split(",");
						for(int m=0; m<strar.length; m++){
							olddata.add(strar[m]);
						}
						String HRID = olddata.get(0);
						if(row==0){
							StringBuilder finaldata = new StringBuilder();
							for(int i=0; i<olddata.size(); i++){
								finaldata.append(olddata.get(i)+",");
							}
							String after_rem = finaldata.substring(0, finaldata.length()-1);
							newdata.add(after_rem);
						}
						
//						System.out.println(oldhr+ " compared to "+HRID);
						if(oldhr.equals(HRID)){
							olddata.set(0, newhr);
							String appendeddata = olddata.get(7);
							olddata.set(2, newhr+"_"+appendeddata);
							olddata.set(6, newhr+"_"+appendeddata+"_NID");
							
							StringBuilder finaldata = new StringBuilder();
							for(int i=0; i<olddata.size(); i++){
								finaldata.append(olddata.get(i)+",");
							}
							String after_rem = finaldata.substring(0, finaldata.length()-1);
							newdata.add(after_rem);
						}
						
						//assigning row
						row++;
					}	
					myInput.close();
			 }
			
			for (int i=0;  i < newdata.size(); i++) {
			}

			//write data to csv file
			FileWriter writer = new FileWriter(new_historycsv);
			for (int i=0;  i < newdata.size(); i++) {
				writer.append(newdata.get(i));
				writer.append('\n');
			}
		    writer.flush();
		    writer.close();
		}
		
		
	}
*/

public void updateHistoryData(String fName) throws IOException{
			ArrayList<String> newdata = new ArrayList<String>();
			int row = 0;		
			
			for(Map.Entry<String,String> map:HRdata.entrySet()){  
				   
				   String oldhr = map.getKey().toString();
				   String newhr = map.getValue().toString();

					String thisLine;
					FileInputStream fis = new FileInputStream(fName);
					DataInputStream myInput = new DataInputStream(fis);
			
					while ((thisLine = myInput.readLine()) != null) {
						
						ArrayList<String> olddata = new ArrayList<String>();
						String strar[] = thisLine.split(",");
						for(int m=0; m<strar.length; m++){
							olddata.add(strar[m]);
						}
						String HRID = olddata.get(0);
						if(row==0){
							StringBuilder finaldata = new StringBuilder();
							for(int i=0; i<olddata.size(); i++){
								finaldata.append(olddata.get(i)+",");
							}
							String after_rem = finaldata.substring(0, finaldata.length()-1);
							newdata.add(after_rem);
						}
						
//						System.out.println(oldhr+ " compared to "+HRID);
						if(oldhr.equals(HRID)){
//							System.out.println("Operation performed!!");
							olddata.set(0, newhr);
							
//							String appendeddata = olddata.get(10);
//							olddata.set(2, newhr+"_"+appendeddata);
							olddata.set(10, newhr);
							
							StringBuilder finaldata = new StringBuilder();
							for(int i=0; i<olddata.size(); i++){
								finaldata.append(olddata.get(i)+",");
							}
							String after_rem = finaldata.substring(0, finaldata.length()-1);
							newdata.add(after_rem);
						}
						
						//assigning row
						row++;
					}	
					myInput.close();
			 }
			/*System.out.println("\n");
			System.out.println("history details ");
			System.out.println(" \n");*/
			
			for (int i=0;  i < newdata.size(); i++) {
				//System.out.println(newdata.get(i));
			}

			//write data to csv file
			FileWriter writer = new FileWriter(new_historycsv);
			for (int i=0;  i < newdata.size(); i++) {
				writer.append(newdata.get(i));
				writer.append('\n');
			}
		    writer.flush();
		    writer.close();
		}
	}
		


