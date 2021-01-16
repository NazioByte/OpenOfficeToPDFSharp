# OpenOfficeToPDFSharp

Convert any Office document file extension to only PDF without Interop.


## Install from Nuget

 

## 🔵 Usage

```csharp

using ZioByte.OpenOffice;

static void Main()
{
   OpenOfficeToPDF officeToPDF = new OpenOfficeToPDF();

   officeToPDF.ConvertStatusChange += OfficeToPDF_ConvertStatusChange;
   officeToPDF.ConvertComplete += OfficeToPDF_ConvertComplete;   

   officeToPDF.Set(@"Simple1.docx", @"\OutputDir\");
   officeToPDF.Set(@"Simple2.docx", @"\OutputDir\");
   officeToPDF.Set(@"Simple3.docx", @"\OutputDir\");
   officeToPDF.Set(@"Simple4.docx", @"\OutputDir\");
   officeToPDF.Set(@"Simple5.docx", @"\OutputDir\");
   
   officeToPDF.Convert();
}

private void OfficeToPDF_ConvertStatusChange(object sender, ConvertStatusChangeEventArgs args)
{
    Console.WriteLine(args.FileName + "," + args.Status);
    // Or do something if you want.
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
   <li>... and other if you want :) </li>  
  </ul>


## Contributing
Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

Please make sure to update tests as appropriate.

## ♥️ License
 [![License](https://img.shields.io/badge/License-Apache%202.0-blue.svg)](https://choosealicense.com/licenses/apache-2.0/)
