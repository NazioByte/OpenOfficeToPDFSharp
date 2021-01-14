# OpenOfficeToPDFSharp

Convert any Office document file extension to only PDF with included LibreOffice 3rd component inside.
 

## 🔵 Usage

```csharp

using ZioByte.OpenOffice;

static void Main()
{
   OpenOfficeToPDF officeToPDF = new OpenOfficeToPDF("demo.docx",  @"\outputDir\");

   officeToPDF.ConvertComplete += OfficeToPDF_ConvertComplete;

   officeToPDF.Convert();
}


private void OfficeToPDF_ConvertComplete(object sender, ConvertCompleteEventArgs args)
{
   //Do something when convert to PDF complete 
}

```

## Contributing
Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

Please make sure to update tests as appropriate.

## ♥️ License
[MIT](https://choosealicense.com/licenses/mit/)
