using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;

namespace ListFieldMigration
{
    class Program
    {
        //propmy user for old column name to copy data from
        public static string GetOldFieldName()
        {
            Console.WriteLine("Enter the Display Name of the old field to copy from:");
            return Console.ReadLine();
        }
        //prompt user for new column name to copy data into
        public static string GetNewFieldName()
        {
            Console.WriteLine("");
            Console.WriteLine("Enter the Display Name of the new field to copy to:");
            return Console.ReadLine();
            
        }
        //
        //create a list of all old/new field pairs for copying
        public static List<KeyValuePair<string, string>> GetFieldsToCopy()
        {
            List<KeyValuePair<string, string>> fieldsList = new List<KeyValuePair<string, string>>();
            bool isDone = false;
            while (isDone == false)
            {
                //prompt the user for the field names, add them to the list
                KeyValuePair<string, string> newPair = new KeyValuePair<string, string>(GetOldFieldName(), GetNewFieldName());
                fieldsList.Add(newPair);
                Console.WriteLine("");
                //propmt the user if they want to add more fields
                Console.WriteLine("Would you like to copy another field?(y/n)");
                string done = Console.ReadLine().ToLower().Trim();
                //if no more fields, set "isDone" to true, which will end the while loop
                if (done == "n" || done == "no")
                {
                    isDone = true;
                }
                else
                {
                    Console.WriteLine("");
                }
            }
            return fieldsList;
        }
        static void Main(string[] args)
        {
            # region Set SPSite
            //User enters site URL in console window
            Console.WriteLine("Enter the URL of the target site:");
            SPSite oSite = new SPSite(Console.ReadLine());
            Console.WriteLine("");

            //Specify a site URL for debugging
            //SPSite oSite = new SPSite(@"http://yourserver");
            # endregion

            # region Set SPWeb
            //User enters web URL in console window
            Console.WriteLine("Enter the path to the target subsite:");
            SPWeb oWeb = oSite.OpenWeb(Console.ReadLine());
            Console.WriteLine("");

            //Specify a web url for debugging
            //SPWeb oWeb = oSite.OpenWeb("/Path/to/target/subsite");
            #endregion

            # region Set SPList
            //User enters list name in console window
            Console.WriteLine("Enter the name of the target list:");
            SPList oList = oWeb.Lists[Console.ReadLine()];
            Console.WriteLine("");

            //Specify a list name for debugging
            //SPList oList = oWeb.Lists["Pages"];
            # endregion

            # region Set copyFields
            //User enters pairs of old and new columns for migration in console window
            List<KeyValuePair<string, string>> copyFields = GetFieldsToCopy();

            //Specify pairs of old and new columns for migration for debugging
            //List<KeyValuePair<string, string>> copyFields = new List<KeyValuePair<string, string>>();
            //copyFields.Add(new KeyValuePair<string,string>("OldColumn1", "NewColumn1"));
            //copyFields.Add(new KeyValuePair<string,string>("OldColumn2", "NewColumn2"));
            # endregion

            int itemCount = 1;
            int totalItems = oList.Items.Count;
            //save state of list properties that need to be changed so they can be restored later
            bool isForcedCheckout = oList.ForceCheckout;
            bool isEnabledModeration = oList.EnableModeration;
            oList.ForceCheckout = false;
            oList.EnableModeration = false;
            oList.Update();
            //create list of files with minor versions that need to be treated specially
            List<SPFile_Info> minorFileInfo = new List<SPFile_Info>();
            try
            {
                foreach (SPFile oFile in oList.RootFolder.Files)
                {
                    try
                    {
                        //if file is checked out, discard the checkout
                        if (oFile.CheckOutType != SPFile.SPCheckOutType.None)
                        {
                            oFile.CheckIn("");
                        }
                        //if file is a minor version or does not have a major version, add its relevent data to our list of minor version files
                        if (oFile.MinorVersion != 0 && oFile.MajorVersion != 0)
                        {
                            SPFile_Info currentFile = new SPFile_Info();
                            currentFile.majorVersion = oFile.MajorVersion;
                            currentFile.minorVersion = oFile.MinorVersion;
                            currentFile.Id = oFile.UniqueId;
                            currentFile.modifiedBy = oFile.Item["Modified By"];
                            currentFile.modifiedDate = oFile.Item["Modified"];
                            minorFileInfo.Add(currentFile);
                        }
                        //if the file is a major version or does not have a major version, migrate the column data and use SystemUpdate(false) to prevent any changes to versioning or modified data
                        else
                        {
                            foreach (KeyValuePair<string, string> oPair in copyFields)
                            {
                                if (oFile.Item[oPair.Key] != null)
                                {
                                    oFile.Item[oPair.Value] = oFile.Item[oPair.Key].ToString();
                                }
                            }
                            oFile.Item.SystemUpdate(false);
                        }
                        Console.WriteLine("Item " + itemCount + " of " + totalItems + " complete.");
                        itemCount++;
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e);
                    }
                }
                Console.WriteLine("");
                Console.WriteLine("Initial Pass Complete.");
                Console.WriteLine("");

                //restore the most recent major version of all files in our minor version list, and publish it to create a new major version
                foreach (SPFile_Info oFile_Info in minorFileInfo)
                {
                    try
                    {
                        SPFile oFile = oWeb.GetFile(oFile_Info.Id);
                        oFile.CheckOut();
                        oFile.Versions.RestoreByLabel(oFile_Info.majorVersion + ".0");
                        oFile.CheckIn("");
                        oFile.Publish("");
                        oFile.Update();
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e);
                    }
                }

                //migrate the column data on the newly restored major version and use SystemUpdate(false) to prevent any changes to versioning or modified data
                foreach (SPFile_Info oFile_Info in minorFileInfo)
                {
                    try
                    {
                        SPListItem oItem = oWeb.GetFile(oFile_Info.Id).Item;
                        foreach (KeyValuePair<string, string> oPair in copyFields)
                        {
                            if (oItem[oPair.Key] != null)
                            {
                                oItem[oPair.Value] = oItem[oPair.Key].ToString();
                            }
                        }
                        oItem["Modified By"] = oFile_Info.modifiedBy.ToString();
                        oItem["Modified"] = oFile_Info.modifiedDate.ToString();
                        oItem.SystemUpdate(false);
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e);
                    }
                }

                //restore the original minor version of each file, and check it in to create a new minor version
                foreach (SPFile_Info oFile_Info in minorFileInfo)
                {
                    try
                    {
                        SPFile oFile = oWeb.GetFile(oFile_Info.Id);
                        oFile.CheckOut();
                        oFile.Versions.RestoreByLabel(oFile_Info.majorVersion + "." + oFile_Info.minorVersion);
                        oFile.CheckIn("");
                        oFile.Update();
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e);
                    }
                }

                //migrate the column data on the newly restored minor version and use SystemUpdate(false) to prevent any changes to versioning or modified data
                foreach (SPFile_Info oFile_Info in minorFileInfo)
                {
                    try
                    {
                        SPListItem oItem = oWeb.GetFile(oFile_Info.Id).Item;
                        foreach (KeyValuePair<string, string> oPair in copyFields)
                        {
                            if (oItem[oPair.Key] != null)
                            {
                                oItem[oPair.Value] = oItem[oPair.Key].ToString();
                            }
                        }
                        oItem["Modified By"] = oFile_Info.modifiedBy.ToString();
                        oItem["Modified"] = oFile_Info.modifiedDate.ToString();
                        oItem.SystemUpdate(false);
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e);
                    }
                }
                Console.WriteLine("Second Pass Complete.");
                Console.WriteLine("");
            }
            finally
            {
                //restore original list property values
                oList.EnableModeration = isEnabledModeration;
                oList.ForceCheckout = isForcedCheckout;
                oList.Update();
            }

            //prompt if user wants to remove old columns from the list
            Console.WriteLine("Do you wish to delete the old columns from the list?(y/n)");
            string deleteFields = Console.ReadLine().ToLower().Trim();
            if (deleteFields == "y" || deleteFields == "yes")
            {
                //remove them
                foreach (KeyValuePair<string, string> oPair in copyFields)
                {
                    oList.Fields[oPair.Key].Delete();
                }
                Console.WriteLine("");
                //prompt if user wants to remove cooresponding site columns as well
                Console.WriteLine("Do you wish to delete the cooresponding Site Columns as well?(y/n)");
                string deleteSiteColumns = Console.ReadLine().ToLower().Trim();
                if (deleteSiteColumns == "y" || deleteSiteColumns == "yes")
                {
                    SPWeb rootWeb = oSite.RootWeb;
                    foreach (KeyValuePair<string, string> oPair in copyFields)
                    {
                        SPField oField = rootWeb.Fields[oPair.Key];
                        //loop through content types and remove site columns so they can be deleted
                        foreach (SPContentType oContentType in rootWeb.ContentTypes)
                        {
                            if (oContentType.Fields.Contains(oField.Id))
                            {
                                oContentType.FieldLinks.Delete(oField.Id);
                                oContentType.Update();
                            }
                        }
                        //delete site columns
                        oField.AllowDeletion = true;
                        oField.Delete();
                    }
                }
            }
            Console.WriteLine("");
            Console.WriteLine("Operation Completed. Press Enter to close this window.");
            Console.ReadLine();
        }
    }
}
