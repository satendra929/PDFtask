package com.notapro.www.pdftask;

import android.content.ActivityNotFoundException;
import android.content.Intent;
import android.content.res.AssetManager;
import android.net.Uri;
import android.os.Environment;
import android.support.v7.app.AppCompatActivity;
import android.os.Bundle;
import android.util.Log;
import android.widget.TextView;
import android.widget.Toast;

import com.itextpdf.text.pdf.PdfReader;
import com.itextpdf.text.pdf.parser.PdfTextExtractor;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.net.URI;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.format.CellFormat;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableCell;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class MainActivity extends AppCompatActivity {
    TextView tv;
    TextView tvDetails;
    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        tv = (TextView) findViewById(R.id.TVDisplay);
        tvDetails = (TextView) findViewById(R.id.TVDetails);
        tv.setText("Candidate: Satendra Varma Email: satvarma@iu.edu Message: Task Complete :)");
        tvDetails.setText("Please find the populated Excel sheet in test folder in internal storage." +
                "Kindly allow pdf scraping to run for 7-10secs before opening the file. ThankYou.");


        //Setting up a writable excel sheet
        AssetManager asm = getAssets();
        WritableWorkbook copy = null;
        try {
            //get original blank copy of the excel sheet in assets folder to make a writable copy
            InputStream ints = asm.open("excel_sheet_cf.xls");
            Workbook wb = Workbook.getWorkbook(ints);
            //making a file to enable writing
            File file = new File(Environment.getExternalStorageDirectory(), "/test");
            if (!file.exists()) {
                file.mkdirs();
            }
            File gpxfile = new File(file, "BeneFix Small Group Plans upload template.xls");
            copy = Workbook.createWorkbook(gpxfile,wb);
        } catch (IOException e) {
            e.printStackTrace();
        } catch (BiffException e) {
            e.printStackTrace();
        }

        //filenames of all the pdfs in the assets folder
        String pdf_names[] = {"para01.pdf","para02.pdf","para03.pdf",
                "para05.pdf","para06.pdf","para07.pdf","para08.pdf","para09.pdf"};

        //variable to keep track of rows in excel sheet
        int rows = 1;
        try {
            for (int p = 0; p<pdf_names.length;p++) {
                PdfReader reader = new PdfReader(getAssets().open(pdf_names[p]));
                int n = reader.getNumberOfPages();
                Log.e("pages",pdf_names[p]);
                //Extracting the content from each page.
                int pages = 1;
                while (pages <= n) {
                    String str = PdfTextExtractor.getTextFromPage(reader, pages);
                    String[] lines = str.split(System.getProperty("line.separator"));
                    parser(lines, copy, rows);
                    rows++;
                    pages++;
                }
                reader.close();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }

        try {
            copy.write();
            copy.close();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (WriteException e) {
            e.printStackTrace();
        }
    }


    //will write the data into the excel sheet(Should have named differently)
    protected void parser(String lines[],WritableWorkbook copy,int row){
        //Setup sheet for writing
        WritableSheet ws = copy.getSheet(0);

        //First line to extract dates
        String dates[] = lines[0].split(" ");
        String start_date = dates[4];
        String end_date = dates[6];
        WritableCell wcs;
        WritableCell wce;
        Label ls = new Label(0, row, start_date);
        Label le = new Label(1, row, end_date);
        wcs = (WritableCell) ls;
        wce = (WritableCell) le;

        //get state abbreviation(get first and last character)
        String state = lines[1].charAt(0)+ "" +lines[1].charAt(lines[1].length()-1); ;
        WritableCell abb;
        Label labb = new Label(3, row, state);
        abb = (WritableCell) labb;

        //get rating area
        String rating_area[] = lines[2].split(" ");
        String area = rating_area[2];
        String numeral = "";
        for (int i =0; i<area.length();i++){
            if(Character.isDigit(area.charAt(i))){
                numeral+= area.charAt(i);
            }
        }
        WritableCell rating_num;
        Label lrn = new Label(4, row, numeral);
        rating_num = (WritableCell) lrn;

        //getting product name
        String product_line[] = lines[3].split(" ");
        String product_name = "";
        for (int i = 5; i <product_line.length;i++){
            product_name = product_name + product_line[i] + " ";
        }
        WritableCell pro_n;
        Label lpn = new Label(2, row, product_name);
        pro_n = (WritableCell) lpn;

        //getting rates for age groups
        String rate5[] = lines[5].split(" ");
        String rateU20 = rate5[1];
        String rate35 = rate5[3];
        String rate50 = rate5[5];
        add_cell(5, row, rateU20, ws);
        add_cell(6, row, rateU20, ws);
        add_cell(21, row, rate35, ws);
        add_cell(36, row, rate50, ws);

        String rate6[] = lines[6].split(" ");
        String rate21 = rate6[1];
        String rate36 = rate6[3];
        String rate51 = rate6[5];
        add_cell(7, row, rate21, ws);
        add_cell(22, row, rate36, ws);
        add_cell(37, row, rate51, ws);

        String rate7[] = lines[7].split(" ");
        String rate22 = rate7[1];
        String rate37 = rate7[3];
        String rate52 = rate7[5];
        add_cell(8, row, rate22, ws);
        add_cell(23, row, rate37, ws);
        add_cell(38, row, rate52, ws);

        String rate8[] = lines[8].split(" ");
        String rate23 = rate8[1];
        String rate38 = rate8[3];
        String rate53 = rate8[5];
        add_cell(9, row, rate23, ws);
        add_cell(24, row, rate38, ws);
        add_cell(39, row, rate53, ws);

        String rate9[] = lines[9].split(" ");
        String rate24 = rate9[1];
        String rate39 = rate9[3];
        String rate54 = rate9[5];
        add_cell(10, row, rate24, ws);
        add_cell(25, row, rate39, ws);
        add_cell(40, row, rate54, ws);

        String rate10[] = lines[10].split(" ");
        String rate25 = rate10[1];
        String rate40 = rate10[3];
        String rate55 = rate10[5];
        add_cell(11, row, rate25, ws);
        add_cell(26, row, rate40, ws);
        add_cell(41, row, rate55, ws);

        String rate11[] = lines[11].split(" ");
        String rate26 = rate11[1];
        String rate41 = rate11[3];
        String rate56 = rate11[5];
        add_cell(12, row, rate26, ws);
        add_cell(27, row, rate41, ws);
        add_cell(42, row, rate56, ws);

        String rate12[] = lines[12].split(" ");
        String rate27 = rate12[1];
        String rate42 = rate12[3];
        String rate57 = rate12[5];
        add_cell(13, row, rate27, ws);
        add_cell(28, row, rate42, ws);
        add_cell(43, row, rate57, ws);

        String rate13[] = lines[13].split(" ");
        String rate28 = rate13[1];
        String rate43 = rate13[3];
        String rate58 = rate13[5];
        add_cell(14, row, rate28, ws);
        add_cell(29, row, rate43, ws);
        add_cell(44, row, rate58, ws);

        String rate14[] = lines[14].split(" ");
        String rate29 = rate14[1];
        String rate44 = rate14[3];
        String rate59 = rate14[5];
        add_cell(15, row, rate29, ws);
        add_cell(30, row, rate44, ws);
        add_cell(45, row, rate59, ws);

        String rate15[] = lines[15].split(" ");
        String rate30 = rate15[1];
        String rate45 = rate15[3];
        String rate60 = rate15[5];
        add_cell(16, row, rate30, ws);
        add_cell(31, row, rate45, ws);
        add_cell(46, row, rate60, ws);

        String rate16[] = lines[16].split(" ");
        String rate31 = rate16[1];
        String rate46 = rate16[3];
        String rate61 = rate16[5];
        add_cell(17, row, rate31, ws);
        add_cell(32, row, rate46, ws);
        add_cell(47, row, rate61, ws);

        String rate17[] = lines[17].split(" ");
        String rate32 = rate17[1];
        String rate47 = rate17[3];
        String rate62 = rate17[5];
        add_cell(18, row, rate32, ws);
        add_cell(33, row, rate47, ws);
        add_cell(48, row, rate62, ws);

        String rate18[] = lines[18].split(" ");
        String rate33 = rate18[1];
        String rate48 = rate18[3];
        String rate63 = rate18[5];
        add_cell(19, row, rate33, ws);
        add_cell(34, row, rate48, ws);
        add_cell(49, row, rate63, ws);

        String rate19[] = lines[19].split(" ");
        String rate34 = rate19[1];
        String rate49 = rate19[3];
        String rate64 = rate19[5];
        add_cell(20, row, rate34, ws);
        add_cell(35, row, rate49, ws);
        add_cell(50, row, rate64, ws);
        add_cell(51, row, rate64, ws);

        try {
            ws.addCell(wcs);
            ws.addCell(wce);
            ws.addCell(pro_n);
            ws.addCell(rating_num);
            ws.addCell(abb);
        } catch (WriteException e) {
            e.printStackTrace();
        }

    }

    public void add_cell(int col,int row,String value, WritableSheet ws ){
        WritableCell cell_n;
        Label lpn = new Label(col, row, value);
        cell_n = (WritableCell) lpn;

        try {
            ws.addCell(cell_n);
        } catch (WriteException e) {
            e.printStackTrace();
        }

    }
}
