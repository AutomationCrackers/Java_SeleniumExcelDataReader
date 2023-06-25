package sel.exceldatareader.Java_SeleniumExcelDataReader;

public class SampleProgram {

	public static void main(String[] args) 
	{
		//Reading the First sheet "Food" in the Excel
		String[][] FoodData = ExcelDataReader.ReadData("TestData","Food");
		for (int i = 0; i < FoodData.length-1; i++) 
		{
			for (int j = 0; j < 2; j++) 
			{
				System.out.println(FoodData[i][j]);
			}
		}
		System.out.println("Food sheet Datas Read Successfully");
		System.out.println("************************************************************");
		//Reading the First sheet "Food" in the Excel
				String[][] PlacesData = ExcelDataReader.ReadData("TestData","Places");
				for (int i = 0; i < PlacesData.length-1; i++) 
				{
					for (int j = 0; j < 2; j++) 
					{
						System.out.println(PlacesData[i][j]);
					}
				}
				System.out.println("Places sheet Datas Read Successfully");
	}

}
