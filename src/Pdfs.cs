using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using SwissAcademic.Citavi;
using SwissAcademic.Citavi.Shell;

public class Pdfs
{
    public static String Export(String path)
    {
        String FOLDERPATH = path + "\\Files\\";
        string Errors = "";
        string Success = "";
        try
        {
            var exportedFilenames = new HashSet<string>();

            ProjectReferenceCollection references = Program.ActiveProjectShell.Project.References;

            foreach (Reference reference in references)
            {
                if (reference.Locations.ToList().Count > 0)
                {
                    foreach (Location loc in reference.Locations)
                    {
                        if (loc.Address.LinkedResourceType.ToString().Equals("NotAnUri") ||
                            loc.Address.LinkedResourceType.ToString().Equals("Empty"))
                        {
                            continue;
                        }
                        else if (loc.LocationType == LocationType.Library)
                        {
                            continue;
                        }

                        var saveFilename = loc.Address.Properties.FileName;

                        if (saveFilename == null)
                        {
                            saveFilename = reference.Id.ToString() + "_" + loc.Index.ToString(); // + ".pdf";
                        }

                        if (exportedFilenames.Contains(saveFilename))
                        {
                            Errors += "[ERROR] Duplicate Filename:" + saveFilename + "\n";
                            continue;
                        }

                        exportedFilenames.Add(saveFilename);
                        var resolvedLocation = "";

                        try
                        {
                            resolvedLocation = loc.Address.Resolve(true).ToString();
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine(e);
                            throw;
                        }


                        //if (loc.Address.LinkedResourceType.ToString().Equals("AttachmentFile"))
                        //{
                        //	try
                        //	{
                        //	    if (loc.Address.ToString().EndsWith(".pdf"))
                        //	    {
                        //	        PdfExportOptions opt = new PdfExportOptions();
                        //	        loc.ExportPdf(FOLDERPATH + saveFilename, opt, true);
                        //	        Success += "Successfully exported local PDF (AttachmentFile): \"" +
                        //	                   reference.Title + "\" Address: \"" + loc.Address.ToString() + "\"\r\n";
                        //	    }
                        //	    else
                        //	    {
                        //	        var test = 0;
                        //	    }
                        //	}
                        //	catch (Exception ex)
                        //	{
                        //		Errors += "[ERROR] Failed to export PDF (AttachmentFile): \"" + reference.Title + "\" Address: \"" + loc.Address.ToString() + "ERROR: "+ex+" "+"\"\r\n";
                        //	}
                        //	
                        //
                        //}
                        //else //if (loc.Address.LinkedResourceType.ToString().Equals("RemoteUri"))
                        //{
                        try
                        {
                            using (WebClient client = new WebClient())
                            {
                                //client.DownloadFile(loc.Address.ToString(), FOLDERPATH + saveFilename);

                                var data = client.DownloadData(resolvedLocation);
                                string extension = "";
                                switch (client.ResponseHeaders["Content-Type"])
                                {
                                    case "application/pdf":
                                    case "application/pdf;charset=Binary":
                                    case "application/pdf;charset=binary":
                                    case "application/pdf;charset=base64":
                                    case "application/pdf;charset=ISO-8859-1":
                                    case "application/pdf; charset=ISO-8859-1":
                                    case "image/pdf;charset=ISO-8859-1":
                                        extension = ".pdf";
                                        break;
                                    case "text/html; charset=utf-8":
                                    case "text/html;charset=UTF-8":
                                    case "text/html":
                                    case "text/html; charset=UTF-8":
                                    case "text/html;charset=utf-8":
                                    case "text/html;charset=ISO-8859-1":
                                    case "text/html; charset=ISO-8859-1":
                                        extension = ".html";
                                        break;
                                    case "image/png":
                                        extension = ".png";
                                        break;
                                    case "image/jpeg":
                                        extension = ".jpg";
                                        break;
                                    case "application/octet-stream":
                                    case null:
                                        extension = "";
                                        break;
                                    default:
                                        Success += "Unknown content type:" +
                                                   client.ResponseHeaders["Content-Type"];
                                        break;
                                }

                                if (!saveFilename.EndsWith(extension))
                                {
                                    saveFilename += extension;
                                }

                                File.WriteAllBytes(FOLDERPATH + saveFilename, data);

                                Success += "Successfully exported remote PDF(RemoteUri): \"" + reference.Title +
                                           "\" Address: \"" + resolvedLocation + "\"\r\n";
                            }
                        }
                        catch (WebException ex)
                        {
                            DebugMacro.WriteLine("Exception while downloading " + reference.Id.ToString());
                            Errors += "[ERROR] Failed to Download RemoteUri (Web Error): \"" + reference.Title +
                                      "\" Address: \"" + resolvedLocation + " Ex: " + ex + "\"\r\n";
                        }

                        //}
                        //else if (loc.Address.LinkedResourceType.ToString().Equals("AbsoluteFileUri"))
                        //{
                        //	try
                        //	{
                        //	    if (loc.Address.ToString().EndsWith(".pdf"))
                        //	    {
                        //	        PdfExportOptions opt = new PdfExportOptions();
                        //	        loc.ExportPdf(FOLDERPATH + saveFilename, opt, true);
                        //	        Success += "Successfully exported PDF (AbsoluteFileUri): \"" + reference.Title +
                        //	                   "\" Address: \"" + loc.Address.ToString() + "\"\r\n";
                        //	    }
                        //	    else
                        //	    {
                        //	        var test = 0;
                        //	    }
                        //    }
                        //	catch (Exception ex)
                        //	{
                        //		Errors += "Failed to export PDF (AbsoluteFileUri): \"" + reference.Title + "\" Address: \"" + loc.Address.ToString() + "ERROR: " + ex + " " + "\"\r\n";
                        //	}
                        //	
                        //	
                        //}
                        //else if (loc.Address.LinkedResourceType.ToString().Equals("AttachmentRemote"))
                        //{
                        //	try
                        //	{
                        //	    //if (loc.Address.ToString().EndsWith(".pdf"))
                        //	    //{
                        //	        PdfExportOptions opt = new PdfExportOptions();
                        //	        ;
                        //	        using (WebClient client = new WebClient())
                        //	        {
                        //	            client.DownloadFile(loc.Address.Resolve(true).ToString(), FOLDERPATH + saveFilename);
                        //	            Success += "Successfully exported remote PDF(RemoteUri): \"" + reference.Title + "\" Address: \"" + loc.Address.ToString() + "\"\r\n";
                        //	        }
                        //
                        //            //loc.ExportPdf(FOLDERPATH + saveFilename, opt, true);
                        //	        //Success += "Successfully exported local PDF (AttachmentRemote): \"" +
                        //	        //           reference.Title + "\" Address: \"" + loc.Address.ToString() + "\"\r\n";
                        //	    //}
                        //	    //else
                        //	    //{
                        //	     //   var test = 0;
                        //	    //}
                        //    }
                        //	catch (Exception ex)
                        //	{
                        //		Errors += "Failed to export PDF (AttachmentRemote): \"" + reference.Title + "\" Address: \"" + loc.Address.ToString() + "ERROR: " + ex + " " + "\"\r\n";
                        //	}
                        //	
                        //	
                        //}
                        //else if (loc.Address.LinkedResourceType.ToString().Equals("NotAnUri"))
                        //{
                        //	Success += loc.Address.ToString() + " NotAnUri \r\n";
                        //}
                        //else if (loc.Address.LinkedResourceType.ToString().Equals("RelativeFileUri"))
                        //{
                        //
                        //	try
                        //	{
                        //	    if (loc.Address.ToString().EndsWith(".pdf"))
                        //	    {
                        //            PdfExportOptions opt = new PdfExportOptions();
                        //		    loc.ExportPdf(FOLDERPATH + saveFilename, opt, true);
                        //		    Success += "Successfully exported local PDF (RelativeFileUri): \"" + reference.Title + "\" Address: \"" + loc.Address.ToString() + "\"\r\n";
                        //	    }
                        //	    else
                        //	    {
                        //	        var test = 0;
                        //	    }
                        //    }
                        //	catch (Exception ex)
                        //	{
                        //		Errors += "Failed to export PDF (RelativeFileUri): \"" + reference.Title + "\" Address: \"" + loc.Address.ToString() + "ERROR: " + ex + " " + "\"\r\n";
                        //	}
                        //	
                        //	
                        //
                        //}
                    }
                }
            }
        }
        catch (Exception e)
        {
            Errors += "Exception: " + e.Message + "\r\n";
            DebugMacro.WriteLine("Exception:" + e.Message);
            DebugMacro.WriteLine(e.StackTrace);

            System.Windows.Forms.MessageBox.Show("Exception:" + e.Message);
        }

        Success += Errors;
        return Success;
    }
}