using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace AsposeTest
{
    public class FileDetails: List<FileDetail>
    {
        public FileDetails(FileInfo[] files)
        {
            foreach (FileInfo file in files.OrderBy(f => f.Length))
            {
                FileDetail fd = new FileDetail(file.FullName, file.Length);
                Add(fd);
            }
        }
    }

    public class FileDetail
    {
        public FileDetail(string name, long length)
        {
            this.FileName = name;
            this.OriginalSize = length;
        }
        public bool Processed { get; set; }
        public string FileName { get; set; }
        public long OriginalSize { get; set; }
        public long OutputSize { get; set; }
    }
}
