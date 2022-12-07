using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using SAPbobsCOM;

namespace ParcelEntry
{
    class ParcelEntry
    {
        private SAPbouiCOM.Application SBO_Application;
        private SAPbobsCOM.Company SBO_Company;
        private SAPbouiCOM.Form oForm;

        #region App Config Data

        string IPath = System.Configuration.ConfigurationManager.AppSettings["IPath"];
        string SaveXML = System.Configuration.ConfigurationManager.AppSettings["SaveXML"];
        int SacCode = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["SacCode"]);
        string Description = System.Configuration.ConfigurationManager.AppSettings["Description"];
        string GSTCODE = System.Configuration.ConfigurationManager.AppSettings["GSTCODE"];
        string ACCCODE = System.Configuration.ConfigurationManager.AppSettings["ACCCODE"];
        #endregion

        public ParcelEntry()
        {
            SetApplication();
            SetConnectionContext();
            singlesignon();
            //AddMenuItems();
            SBO_Application.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(SBO_Application_MenuEvent);
            SBO_Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);            
        }

        #region SAP B1 Application Connection
        private void SetApplication()
        {
            try
            {
                SBO_Company = new SAPbobsCOM.Company();
                SAPbouiCOM.SboGuiApi SboGuiApi = null;
                string sConnectionString = null;
                SboGuiApi = new SAPbouiCOM.SboGuiApi();
                sConnectionString = System.Convert.ToString(Environment.GetCommandLineArgs().GetValue(1));
                SboGuiApi.Connect(sConnectionString);   //If there's no active application the connection will fail
                SBO_Application = SboGuiApi.GetApplication(-1);
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox(ex.ToString());
            }
        }
        private int SetConnectionContext()
        {
            int setConnectionContextReturn = 0;
            string sCookie = null;
            string sConnectionContext = null;
            SBO_Company = new SAPbobsCOM.Company();
            sCookie = SBO_Company.GetContextCookie();
            sConnectionContext = SBO_Application.Company.GetConnectionContext(sCookie);
            if (SBO_Company.Connected == true)
            {
                SBO_Company.Disconnect();
            }
            setConnectionContextReturn = SBO_Company.SetSboLoginContext(sConnectionContext);
            return setConnectionContextReturn;
        }
        private void singlesignon()
        {
            try
            {
                SBO_Company = SBO_Application.Company.GetDICompany();
                SBO_Application.StatusBar.SetText("Welcome to " + " " + SBO_Company.CompanyName, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox(ex.ToString());
            }
        }
        #endregion

        #region Add PTPL & Parcel Entry Menu in SAP B1 Module
        private void AddMenuItems()
        {
            SAPbouiCOM.Menus oMenus = null;
            SAPbouiCOM.MenuItem oMenuItem = null;
            SAPbouiCOM.MenuCreationParams oCreationPackage = null;
            oMenus = SBO_Application.Menus;  //Get the menus collection from the application
            oCreationPackage = ((SAPbouiCOM.MenuCreationParams)(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)));
            oMenuItem = SBO_Application.Menus.Item("43520");  //Moudles
            try
            {
                if (oMenus.Exists("PTPL"))
                {
                    oMenus.RemoveEx("PTPL");
                }
                oMenuItem = SBO_Application.Menus.Item("43520");
                oMenus = oMenuItem.SubMenus;
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                oCreationPackage.UniqueID = "PTPL";
                oCreationPackage.String = "Priyanka Traders";
                oCreationPackage.Enabled = true;
                oCreationPackage.Image = IPath;
                oCreationPackage.Position = 16;
                oMenus.AddEx(oCreationPackage);
                oMenuItem = SBO_Application.Menus.Item("PTPL");  // Get the menu collection of the newly added pop-up item
                oMenus = oMenuItem.SubMenus;
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;    // Create a sub menu
                try
                {
                    oCreationPackage.UniqueID = "ParcelEntry";
                    oCreationPackage.String = "Parcel Entry";
                    oCreationPackage.Position = 0;
                    oMenus.AddEx(oCreationPackage);
                }
                catch (Exception ex)
                {
                    SBO_Application.MessageBox(ex.Message);
                }
            }
            catch (Exception ex)
            {
                ex.Message.ToString();
            }
        }
        #endregion

        private void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            //string getmenuid = pVal.MenuUID;
            BubbleEvent = true;
            try
            {
                #region Load Form On SBO On Parcel Entry Menu Click

                if ((pVal.MenuUID == "ParcelEntry") && (pVal.BeforeAction == false))
                {
                    string FormName = @"ParcelEntry.srf";
                    LoadFromXML(FormName);
                    oForm = SBO_Application.Forms.Item("P_Entry");
                    oForm.EnableMenu("1288", true);
                    oForm.EnableMenu("1289", true);
                    oForm.EnableMenu("1290", true);
                    oForm.EnableMenu("1291", true);
                    oForm.Visible = true;
                }
                #endregion

                #region Add Document Number Automatically On Form

                if (pVal.MenuUID == "1282")
                {
                    SAPbouiCOM.EditText oDocNum = oForm.Items.Item("DocNum").Specific;
                    oForm.Items.Item("DocNum").Enabled = false;
                    SAPbobsCOM.Recordset orec = null;
                    orec = ((SAPbobsCOM.Recordset)(SBO_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                    orec.DoQuery("Select DocEntry From [@PARCELENTRY]");
                    int chk = orec.RecordCount;
                    oDocNum.Value = (chk + 1).ToString();

                    //SAPbouiCOM.EditText oDate = oForm.Items.Item("Date").Specific;
                    //string date = DateTime.Now.ToString("ddMMyyyy");
                    //oDate.Value = date;
                    //oDate.Value = (DateTime.Now).ToString();
                }
                #endregion


                #region Enable Arrow on SAP Form Not Working Yet

                //1281 find button


                //SAPbobsCOM.Recordset oorec = null;
                //oorec = ((SAPbobsCOM.Recordset)(SBO_Company.GetBusinessObject(SAPbobsCOM.UserPermissionForms.Equals())));
                //SAPbobsCOM.UserPermissionTree mUserPermission;
                //mUserPermission = SBO_Company.GetBusinessObject(oUserPermissionTree);
                //string p = pVal.MenuUID;
                //switch (p)
                //{
                //    case "1288":
                //        oorec.MoveNext();
                //        break;
                //    case "1289":
                //        oorec.MovePrevious();
                //        break;
                //    case "1290":
                //        oorec.MoveFirst();
                //        break;
                //    case "1291":
                //        oorec.MoveLast();
                //        break;
                //}
                #endregion

            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox(ex.Message);
            }
        }

        private void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            #region Select Only Require Field from CFL

            if (pVal.FormUID == "P_Entry" && pVal.ItemUID == "SCode")
            {
                AddChooseFromListRule("CFLOCRD", "CardType", "S");      //Get Only Supplier Code in Supplier Code Field 
            }
            if (pVal.FormUID == "P_Entry" && pVal.ItemUID == "TransCode")
            {
                AddChooseFromListRule("CFLOCRD2", "groupcode", "104");  //Get Only Transporter Code in Transporter Code
            }
            #endregion

            #region Fill Name Field Using CFL Code

            if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
            {
                SAPbouiCOM.Form oForm = null;
                oForm = SBO_Application.Forms.Item(FormUID);
                SAPbouiCOM.IChooseFromListEvent oCFLEvento = null;
                oCFLEvento = ((SAPbouiCOM.IChooseFromListEvent)(pVal));
                string sCFL_ID = null;
                sCFL_ID = oCFLEvento.ChooseFromListUID;
                SAPbouiCOM.ChooseFromList oCFL = null;
                oCFL = oForm.ChooseFromLists.Item(sCFL_ID);
                string code = null;
                string name = null;

                if (sCFL_ID == "CFLOCRD")
                {
                    if (oCFLEvento.BeforeAction == false)
                    {
                        SAPbouiCOM.DataTable oDataTable = null;
                        oDataTable = oCFLEvento.SelectedObjects;                       
                        try
                        {
                            code = System.Convert.ToString(oDataTable.GetValue(0, 0));
                            name = System.Convert.ToString(oDataTable.GetValue(1, 0));
                        }
                        catch (Exception)
                        {
                            throw;
                        }
                        if ((pVal.ItemUID == "SCode"))
                        {
                            oForm = SBO_Application.Forms.Item("P_Entry");
                            SAPbouiCOM.EditText oSuppName = oForm.Items.Item("SName").Specific;
                            oSuppName.Value = name;
                        }
                    }
                }

                else if (sCFL_ID == "CFLOCRD2")
                {
                    if (oCFLEvento.BeforeAction == false)
                    {
                        SAPbouiCOM.DataTable oDataTable = null;
                        oDataTable = oCFLEvento.SelectedObjects;                       
                        try
                        {
                            code = System.Convert.ToString(oDataTable.GetValue(0, 0));
                            name = System.Convert.ToString(oDataTable.GetValue(1, 0));
                        }
                        catch (Exception)
                        {
                            throw;
                        }
                        if ((pVal.ItemUID == "TransCode"))
                        {
                            oForm = SBO_Application.Forms.Item("P_Entry");
                            SAPbouiCOM.EditText oSuppName = oForm.Items.Item("TransName").Specific;
                            oSuppName.Value = name;
                        }
                    }
                }

            }
            #endregion

            #region Create Invoice IF Pay By PTPL and Pay Type Immediate

            if (pVal.FormUID == "P_Entry" && pVal.ItemUID == "1")
            {
                string PayBy = null;
                string PayType = null;
                oForm = SBO_Application.Forms.Item("P_Entry");
                SAPbouiCOM.ComboBox oPaidBy = oForm.Items.Item("TCPB").Specific;
                SAPbouiCOM.ComboBox oPmtType = oForm.Items.Item("PayType").Specific;
                PayBy = oPaidBy.Value;
                PayType = oPmtType.Value;

                string tCode = null;
                SAPbouiCOM.EditText oSuppName = oForm.Items.Item("TransCode").Specific;
                tCode = oSuppName.Value;

                string DocNo = null;
                SAPbouiCOM.EditText oDocNo = oForm.Items.Item("DocNum").Specific;
                DocNo = oDocNo.Value;

                string AmtValue = null;
                SAPbouiCOM.EditText oAmtValue = oForm.Items.Item("TransCharg").Specific;
                AmtValue = oAmtValue.Value;

                if (PayBy == "P" && PayType == "I")
                {
                    ADDAPINVOICE(tCode, DocNo, AmtValue);
                }
            }
            #endregion

        }
        

        #region Add A/P Invoice
        public void ADDAPINVOICE(string TransPorterCode,string DocNo ,string AmtValue)
        {
            try
            {
                SAPbobsCOM.Documents oPInvoice;
                oPInvoice = SBO_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices);
                oPInvoice.CardCode = TransPorterCode;
                oPInvoice.DocDate = DateTime.Now;
                oPInvoice.DocDueDate = DateTime.Now;
                oPInvoice.BPL_IDAssignedToInvoice = 1;
                oPInvoice.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service;
                oPInvoice.NumAtCard = DocNo;                    //Parcel Entry Document Number


                //oPInvoice.ShipState = "MH";
                //oPInvoice.ShipPlace = "MH";
                //oPInvoice.ShipFrom = "MH";
                //oPInvoice.ShipToCode = "MH";
                //oPInvoice.GSTTransactionType = SAPbobsCOM.GSTTransactionTypeEnum.gsttrantyp_BillOfSupply;                
                //oPInvoice.AddressExtension.PurchasePlaceOfSupply = "MH";
                //oPInvoice.AddressExtension.DeliveryPlaceState = "MH";
                //oPInvoice.ShipFrom = "Ship To";


                oPInvoice.Lines.SetCurrentLine(0);
                oPInvoice.Lines.SACEntry = SacCode;             // From OSAC table
                oPInvoice.Lines.ItemDescription = Description;
                oPInvoice.Lines.TaxCode = GSTCODE;
                oPInvoice.Lines.AccountCode = ACCCODE;
                oPInvoice.Lines.Quantity = 1;
                oPInvoice.Lines.UnitPrice = Convert.ToInt32(AmtValue);
                oPInvoice.Lines.LocationCode = 2;

                if (oPInvoice.Add() != 0)
                {
                    int irrcode;
                    string errmsg;
                    SBO_Company.GetLastError(out irrcode, out errmsg);
                    SBO_Application.StatusBar.SetText(errmsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    SBO_Application.StatusBar.SetText("Invoice Created Successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                }
            }
            catch (Exception ex)
            {
                ex.ToString();
            }
        }
        #endregion

        #region Load B1S/SRF Form in Application (Parcel Entry Form in SAP B1)
        private void LoadFromXML(string FileName)
        {
            try
            {
                XmlDocument oXmlDoc = new XmlDocument();
                string sPath = Application.StartupPath.ToString();
                oXmlDoc.Load(sPath + @"\" + FileName);      //Load as XML Document           
                string strXML = oXmlDoc.InnerXml.ToString();
                SBO_Application.LoadBatchActions(ref strXML);
                oXmlDoc.Save(SaveXML);                      //Save as XML
            }
            catch (Exception ex)
            {
                ex.ToString();
            }
        }
        #endregion       

        #region Add Rule in CFL (Take only Vendor and Transloader List in CFL)
        private void AddChooseFromListRule(string CFLID, string alias, string Condition)
        {
            try
            {
                SAPbouiCOM.Conditions oConditions;
                SAPbouiCOM.Condition oCondition;
                SAPbouiCOM.ChooseFromList oChooseFromList;
                SAPbouiCOM.Conditions emptyCon = null;
                oChooseFromList = SBO_Application.Forms.Item("P_Entry").ChooseFromLists.Item(CFLID);
                oChooseFromList.SetConditions(emptyCon);
                oConditions = oChooseFromList.GetConditions();
                oCondition = oConditions.Add();
                oCondition.Alias = alias;
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCondition.CondVal = Condition;
                oChooseFromList.SetConditions(oConditions);
            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText(ex.Message);
            }
        }
        #endregion 
    }
}