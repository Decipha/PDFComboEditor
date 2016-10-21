using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Pdf;

namespace PDFTest
{
	class Class1
	{
		public static void Test()
		{
			// Create a PDF license object
			Aspose.Pdf.License license = new Aspose.Pdf.License();
			// Instantiate license file
			
			// Set the value to indicate that license will be embedded in the application
			license.Embedded = true;

		}


	}
}
