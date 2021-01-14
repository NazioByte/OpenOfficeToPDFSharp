# OpenOfficeToPDFSharp

Convert any Office document file extension to only PDF.

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
   //Do something when convert PDF to complete.
}

```

## File Extension Support

 <ul>
  <li>.doc</li>
<li>.docx</li>
<li>.txt</li>
<li>.rtf</li>
<li>.html</li>
<li>.htm</li>
<li>.xml</li>
<li>.wps</li>
<li>.wpd</li>
<li>.xls</li>
<li>.xlsb</li>
<li>.xlsx</li>
<li>.ods</li>
<li>.ppt</li>
<li>.pptx</li>
<li>.odp</li>  
  </ul>


## Contributing
Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

Please make sure to update tests as appropriate.

## ♥️ License
 [![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
