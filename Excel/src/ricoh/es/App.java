package ricoh.es;

import java.util.ArrayList;

public class App {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		versionV1();
	}
	private static void versionV1(){
		String pathOrigin = "C:\\Users\\Agustí\\Desktop\\Ma_documents\\CVs\\RICOH\\DemoOrigin.xlsm";
		String pathDestin = "C:\\Users\\Agustí\\Desktop\\Ma_documents\\CVs\\RICOH\\DemoDestin.xlsm";
		
		ricoh.es.methods.Excel excel = new ricoh.es.methods.Excel(pathOrigin, pathDestin);
		String HeaderX = excel.readCell("Hoja1", 0, 0);
		String HeaderY = excel.readCell("Hoja1", 0, 1);
		//excel.createSheet("Fulla1");
		excel.writeCell(0, 0, HeaderX , "Hoja1");
		excel.writeCell(0, 1, HeaderY , "Hoja1");
		
		int max = 0;		


		ArrayList<Object[]> al = new ArrayList<Object[]>();
		Object[] ob;
		while(excel.readCell("Hoja1", ++max, 0) != null){
			ob = new Object[2];
			ob[0] = excel.readCell("Hoja1", max, 0);
			ob[1] = excel.readCell("Hoja1", max, 1);	
			al.add(ob);
		}
		for(int i=0;i<al.size();i++){	
			//System.out.println(al.get(i)[1]);
		}
		
		int n = 0;
		for(int i=5;i<15;i++){
			excel.writeCell(++n, 0, Double.parseDouble(al.get(i)[0].toString()) , "Hoja1");
			excel.writeCell(n, 1,  Double.parseDouble(al.get(i)[1].toString()) , "Hoja1");
		}
		
		ricoh.es.methods.Utils.openFile(pathDestin);
		
	}

}
