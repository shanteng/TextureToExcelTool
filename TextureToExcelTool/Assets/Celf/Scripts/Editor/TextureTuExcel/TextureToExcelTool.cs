#if UNITY_EDITOR
using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using UnityEditor;
using System.Text.RegularExpressions;
using System.IO;
using System.Linq;
using System.Drawing;
using System;
using OfficeOpenXml;
using OfficeOpenXml.Style;

public static class TextureToExcelTool
{
    [MenuItem("Assets/导出图片到Excel", false, 10)]
    static private void ExportTextureToExcel()
    {
       
        List<string> fullList = new List<string>();
        for (int iindex = 0; iindex < Selection.objects.Length; ++iindex)
        {
            string full = AssetDatabase.GetAssetPath(Selection.objects[iindex]);
            if (AssetDatabase.IsValidFolder(full))
            {
                DirectoryInfo directory = new DirectoryInfo(full);
                FileInfo[] files = directory.GetFiles("*", SearchOption.AllDirectories);//查找改路径下的所有文件夹，包含子文件夹
                for (int i = 0; i < files.Length; i++)
                {
                    bool isTexture = files[i].Name.EndsWith(".png") || files[i].Name.EndsWith(".jpg");
                    if (!isTexture)
                        continue;
                    string m_CurrentTexturePath = files[i].FullName.Substring(files[i].FullName.IndexOf(@"Assets\"), files[i].FullName.Length - files[i].FullName.IndexOf(@"Assets\"));
                    fullList.Add(m_CurrentTexturePath);
                }
            }
            else
            {
                fullList.Add(full);
            }
        }

        string excelPath = System.Environment.CurrentDirectory + "/Assets/Celf/Scripts/Editor/TextureTuExcel/";


        FileInfo existingFile = new FileInfo(excelPath+ "UiTexture.xlsx");
        ExcelPackage _package = new ExcelPackage(existingFile);
        ExcelWorksheet _worksheet = _package.Workbook.Worksheets["LocalizationText"];


        if (existingFile.Exists)
        {
            if (System.IO.File.GetAttributes(excelPath).ToString().IndexOf("ReadOnly") != -1)
                File.SetAttributes(excelPath, FileAttributes.Normal);
             _package = new ExcelPackage(existingFile);
             _worksheet = _package.Workbook.Worksheets["LocalizationText"];
        }
        else
        {
            _package = new ExcelPackage();
            _worksheet = _package.Workbook.Worksheets.Add("LocalizationText");
        }

        int imgCol = 3;
        _worksheet.Cells[1, 1].Value = "ID";
        _worksheet.Cells[1, 2].Value = "Name";
        _worksheet.Cells[1, imgCol].Value = "Image---------------------------------------------------------------------------------------------------------------------------------------------";

        int _curRows = _worksheet.Dimension.Rows;

        for (int i = 0; i < fullList.Count; i++)
        {
            string m_CurrentTexturePath = fullList[i];
            int row = _curRows + 1;
            _worksheet.Cells[row, 1].Value = _curRows;
            _worksheet.Cells[row, 2].Value = m_CurrentTexturePath;
            _worksheet.Cells[row, imgCol].Style.Fill.PatternType = ExcelFillStyle.Solid;
            _worksheet.Cells[row, imgCol].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 0, 0));//设置单元格背景色为黑色

            for (int col = 1; col <= 5; ++col)
            {
                _worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                _worksheet.Cells[row, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;//垂直居中
            }

            using (var image = Image.FromFile(m_CurrentTexturePath))
            {
                double MinHeight = 100;
                double imgeWidht = MinHeight / image.Height * image.Width;

                var picture = _worksheet.Drawings.AddPicture($"image_{DateTime.Now.Ticks}", image);
                picture.SetPosition(row - 1, 20, imgCol - 1, 0);
                picture.SetSize((int)imgeWidht, (int)MinHeight);

                _worksheet.Cells.Style.ShrinkToFit = true;//单元格自动适应大小
                _worksheet.Row(row).Height = MinHeight;//设置行高
                _worksheet.Row(row).CustomHeight = true;//自动调整行高
            }

            _curRows++;
        }

        FileOutputUtil.OutputDir = new DirectoryInfo(excelPath);
        var xlFile = FileOutputUtil.GetFileInfo("UiTexture.xlsx");
        _worksheet.Cells.AutoFitColumns(0);
        _package.SaveAs(xlFile);

        bool result =  EditorUtility.DisplayDialog("","Fininshed","ok");
        if (result)
        {
            EditorUtility.RevealInFinder(excelPath + "UiTexture.xlsx");
        }
    }


}

#endif