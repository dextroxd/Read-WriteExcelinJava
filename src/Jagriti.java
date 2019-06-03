import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Scanner;

public class Jagriti
{

    public static void main(String args[])
    {
        File file = new File("C:\\Users\\Divyanshu Dhawan\\Desktop\\Boxes.xls");
        File file1 = new File("C:\\Users\\Divyanshu Dhawan\\Desktop\\Products.xls");
        ArrayList<String> boxes = new ArrayList<>();
        ArrayList<String> products = new ArrayList<>();
        try {
            Workbook workbook = Workbook.getWorkbook(file);
            Workbook workbook1 = Workbook.getWorkbook(file1);
            Sheet sheet = workbook.getSheet(0);
            Sheet sheet1 = workbook1.getSheet(0);
//            Cell cell = sheet.getCell(0,0);
//            System.out.println(cell.getContents());
            for(int i=0;i<sheet.getRows();i++){
                Cell cell = sheet.getCell(0,i);
                boxes.add(cell.getContents());
            }
            for(int i=0;i<sheet1.getRows();i++){
                Cell cell = sheet1.getCell(0,i);
                products.add(cell.getContents());
            }
        } catch (IOException e) {
            e.printStackTrace();
        } catch (BiffException e) {
            e.printStackTrace();
        }
//    for(int i=0;i<products.size();i++){
//            System.out.println(products.get(i));
//    }
        if(!boxes.isEmpty()&&!products.isEmpty()){
            try {
                System.out.println("Enter no. of product");
                Scanner sc = new Scanner(System.in);
                int n = sc.nextInt();
                int count = 0;
                WritableWorkbook writableWorkbook = Workbook.createWorkbook(new File("C:\\Users\\Divyanshu Dhawan\\Desktop\\Final.xls"));
                writableWorkbook.createSheet("firstSheet",0);
                WritableSheet copySheet = writableWorkbook.getSheet(0);
                for(int box=0;box<boxes.size()+products.size();box+=1){
                    if(box%(n+1)==0){
                        count +=1;
                        Label label = new Label(0,box,boxes.get(box/(n+1)));
                        copySheet.addCell(label);


                    }
                    else{
                        Label label = new Label(0,box,products.get(box-count));
                        copySheet.addCell(label);

                    }
                }
                writableWorkbook.write();
                writableWorkbook.close();

            } catch (IOException e) {
                e.printStackTrace();
            } catch (RowsExceededException e) {
                e.printStackTrace();
            } catch (WriteException e) {
                e.printStackTrace();
            }
        }

    }


}
