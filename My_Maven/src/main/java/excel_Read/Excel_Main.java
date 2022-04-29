package excel_Read;

import java.io.IOException;

public class Excel_Main {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		
		String s=Excel_Code.readStringData(0, 0); // string data is in 0th row and 0th column in sheet
		System.out.println(s);
		System.out.println();
		String s1=Excel_Code.readIntegerData(3, 0); // integer data in 3rd row and 0th column
		System.out.println(s1);
		

	}

}
