# PrintOrder_WPF_VSTO

This module is designed to create an order at the car service.
Steps of work:
1) In the form, the data is entered with the order number, the necessary data and the choice of the template
2) Scan template form from excel and create clone;
3) Adding data from database to template clone;
4) Output the clone of the completed template for further action (view, save, print, etc.).

The following technologies were used to create this module:
 - WPF (XAML);
 - CSVSTOViewEXcelInWPF (using Microsoft.Office.Interop.Excel);
 - MySQL;
 - MaterialDesignToolkit.Wpf;
 - BarcodeLib (generator barcode);
 - Gma.QrCodeNet.Encoding (generator qrcode).
 
 Need the next packages (with version do not least):
 - linq2db.1.8.3
 - linq2db.MySql.1.8.3
 - linq2db.t4models.1.8.3
 - MaterialDesignColors.1.1.3
 - MaterialDesignThemes.2.3.0.823
