﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

public class FileOutputUtil
{
    static DirectoryInfo _outputDir = null;
    public static DirectoryInfo OutputDir
    {
        get
        {
            return _outputDir;
        }
        set
        {
            _outputDir = value;
            if (!_outputDir.Exists)
            {
                _outputDir.Create();
            }
        }
    }
    public static FileInfo GetFileInfo(string file, bool deleteIfExists = true)
    {
        var path = OutputDir.FullName + Path.DirectorySeparatorChar + file;
        var fi = new FileInfo(path);
        if (fi.Exists)
        {
            if (System.IO.File.GetAttributes(path).ToString().IndexOf("ReadOnly") != -1)
            {
                File.SetAttributes(path, FileAttributes.Normal);
            }
        }
        if (deleteIfExists && fi.Exists)
        {
            fi.Delete();  // ensures we create a new workbook
        }
        return fi;
    }
    public static FileInfo GetFileInfo(DirectoryInfo altOutputDir, string file, bool deleteIfExists = true)
    {
        var fi = new FileInfo(altOutputDir.FullName + Path.DirectorySeparatorChar + file);
        if (deleteIfExists && fi.Exists)
        {
            fi.Delete();  // ensures we create a new workbook
        }
        return fi;
    }


    internal static DirectoryInfo GetDirectoryInfo(string directory)
    {
        var di = new DirectoryInfo(_outputDir.FullName + Path.DirectorySeparatorChar + directory);
        if (!di.Exists)
        {
            di.Create();
        }
        return di;
    }
}
