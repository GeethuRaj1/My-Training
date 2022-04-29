package excel_Read_Instance;

import java.io.IOException;

public class Excel_Instance_Main {

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		Excel_Instance_Code ob=new Excel_Instance_Code();
		String s=ob.readSheet(0, 0);
		System.out.println("String value is:"+s);
		String s1=ob.readSheet(3, 0);
		System.out.println("Integer value is:"+s1);
				
	}

}
