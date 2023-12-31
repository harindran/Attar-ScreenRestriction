﻿using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ScreenRestriction.Common
{
    class clsMenuEvent
    {
        SAPbouiCOM.Form objform;
        string strsql;
        public void MenuEvent_For_StandardMenu(ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                switch (clsModule. objaddon.objapplication.Forms.ActiveForm.TypeEx)
                {
                    case "133":
                    case "-392":
                        {
                            // Default_Sample_MenuEvent(pVal, BubbleEvent)
                            if (pVal.BeforeAction == true)
                                return;
                            objform = clsModule.objaddon.objapplication.Forms.ActiveForm;
                            Default_Sample_MenuEvent(pVal, BubbleEvent);

                            break;
                        }
                   
                    case "USRSCRN":
                        UserConfiguration_MenuEvent(ref pVal, ref BubbleEvent);
                        break;
                }
            } 
            catch (Exception ex)
            {

            }
        }

        private void Default_Sample_MenuEvent(SAPbouiCOM.MenuEvent pval, bool BubbleEvent)
        {
            try
            {
                objform = clsModule.objaddon.objapplication.Forms.ActiveForm;
                if (pval.BeforeAction == true)
                {
                }

                else
                {
                    SAPbouiCOM.Form oUDFForm;
                    try
                    {
                        oUDFForm = clsModule.objaddon.objapplication.Forms.Item(objform.UDFFormUID);
                    }
                    catch (Exception ex)
                    {
                        oUDFForm = objform;
                    }

                    switch (pval.MenuUID)
                    {
                        case "1281": // Find
                            {
                                //oUDFForm.Items.Item("U_RevRecDN").Enabled = true;
                                break;
                            }
                        case "1287":
                            {
                                //if (oUDFForm.Items.Item("U_RevRecDN").Enabled == false)
                                //{
                                //    oUDFForm.Items.Item("U_RevRecDN").Enabled = true;
                                //}
                                //((SAPbouiCOM.EditText)oUDFForm.Items.Item("U_RevRecDN").Specific).String = "";
                                break;
                            }
                        default:
                            {
                                //oUDFForm.Items.Item("U_RevRecDN").Enabled = false;
                                break;
                            }
                    }
                }
            }
            catch (Exception ex)
            {
                // objaddon.objapplication.SetStatusBarMessage("Error in Standart Menu Event" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            }
        }
              

        private void UserConfiguration_MenuEvent(ref SAPbouiCOM.MenuEvent pval, ref bool BubbleEvent)
        {
            try
            {
                SAPbouiCOM.DBDataSource DBSource;
                SAPbouiCOM.Matrix Matrix0;
                
                objform = clsModule.objaddon.objapplication.Forms.ActiveForm;
                DBSource = objform.DataSources.DBDataSources.Item("@AT_USRSCRN");
                Matrix0 = (SAPbouiCOM.Matrix)objform.Items.Item("mtxlist").Specific;
                if (pval.BeforeAction == true)
                {
                    switch (pval.MenuUID)
                    {
                        case "1283":
                            if (clsModule.objaddon.objapplication.MessageBox("Removing of an entry cannot be reversed. Do you want to continue?", 2, "Yes", "No") != 1)
                            {
                                BubbleEvent = false;
                            }
                            break;
                        case "1293":
                            if (Matrix0.VisualRowCount == 1) BubbleEvent = false;
                            break;
                    }
                }
                else
                {
                    switch (pval.MenuUID)
                    {
                        case "1281": // Find Mode                            
                            objform.Items.Item("tuserid").Enabled = true;
                            objform.Items.Item("tname").Enabled = true;
                            Matrix0.Item.Enabled = false;   
                            break;
                        case "1293"://Delete Row
                            DeleteRow(Matrix0, "@AT_USRSCRN1");
                            break;
                        case "1282"://Add Mode                            
                            ((SAPbouiCOM.EditText)objform.Items.Item("trem").Specific).String = "Created By " + clsModule.objaddon.objcompany.UserName + " on " + DateTime.Now.ToString("dd/MMM/yyyy HH:mm:ss");                           
                            break;                       

                    }
                }
            }
            catch (Exception ex)
            {

            }
        }

        public void DeleteRow(SAPbouiCOM.Matrix objMatrix, string TableName)
        {
            try
            {
                SAPbouiCOM.DBDataSource DBSource;
                // objMatrix = objform.Items.Item("20").Specific
                objMatrix.FlushToDataSource();
                DBSource = objform.DataSources.DBDataSources.Item(TableName); 
                for (int i = 1, loopTo = objMatrix.VisualRowCount; i <= loopTo; i++)
                {
                    objMatrix.GetLineData(i);
                    DBSource.Offset = i - 1;
                    DBSource.SetValue("LineId", DBSource.Offset, Convert.ToString(i));
                    objMatrix.SetLineData(i);
                    objMatrix.FlushToDataSource();
                }
                DBSource.RemoveRecord(DBSource.Size - 1);
                objMatrix.LoadFromDataSource();
            }

            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText("DeleteRow  Method Failed:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None);
            }
            finally
            {
            }
        }

        private bool Cancelling_IntBranch_RecoJournalEntry(string FormUID, string JETransId)
        {            
                string TransId;
                SAPbouiCOM.Matrix objmatrix;
                SAPbobsCOM.JournalEntries objjournalentry;
                if (string.IsNullOrEmpty(JETransId))
                    return true;
                SAPbobsCOM.Recordset objRs;
                string strSQL;
                try
                {
                    objform = clsModule.objaddon.objapplication.Forms.Item(FormUID);
                    objmatrix =(SAPbouiCOM.Matrix) objform.Items.Item("mtxcont").Specific;
                    objRs = (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                    string GetStatus = clsModule.objaddon.objglobalmethods.getSingleValue("select distinct 1 as \"Status\" from OJDT where \"StornoToTr\"=" + JETransId + "");
                    if (GetStatus == "1")
                    {
                        TransId = clsModule.objaddon.objglobalmethods.getSingleValue("select \"TransId\" from OJDT where \"StornoToTr\"=" + JETransId + "");
                        ((SAPbouiCOM.EditText)objform.Items.Item("trvtran").Specific).String= TransId;
                    //return true;
                    }
                    strSQL = "Select T0.\"Series\",T0.\"TaxDate\",T0.\"DueDate\",T0.\"RefDate\",T0.\"Ref1\",T0.\"Ref2\",T0.\"Memo\",T1.\"Account\",T1.\"Credit\",T1.\"Debit\",T1.\"BPLId\",T0.\"U_RevRecDN\",T0.\"U_RevRecDE\",T1.\"U_InvEntry\",";
                    strSQL += "\n (Select \"CardCode\" from OCRD where \"CardCode\"=T1.\"ShortName\") as \"BPCode\"";
                    strSQL += "\n from OJDT T0 join JDT1 T1 ON T0.\"TransId\"=T1.\"TransId\" where  T1.\"TransId\"='" + JETransId + "' order by T1.\"Line_ID\"";
                    objRs.DoQuery(strSQL);
                    if (objRs.RecordCount == 0)
                        return true;
                    if (!clsModule.objaddon.objcompany.InTransaction) clsModule.objaddon.objcompany.StartTransaction();
                    objjournalentry = (SAPbobsCOM.JournalEntries)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
                    clsModule.objaddon.objapplication.StatusBar.SetText("Journal Entry Reversing Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                    objjournalentry.TaxDate = Convert.ToDateTime(objRs.Fields.Item("TaxDate").Value); // objJEHeader.GetValue("TaxDate", 0)
                    objjournalentry.DueDate = Convert.ToDateTime(objRs.Fields.Item("DueDate").Value); // objJEHeader.GetValue("DueDate", 0)
                    objjournalentry.ReferenceDate = Convert.ToDateTime(objRs.Fields.Item("RefDate").Value); // objJEHeader.GetValue("RefDate", 0)
                    objjournalentry.Reference = Convert.ToString(objRs.Fields.Item("Ref1").Value); // objJEHeader.GetValue("Ref1", 0)
                    objjournalentry.Reference2 = Convert.ToString(objRs.Fields.Item("Ref2").Value); // objJEHeader.GetValue("Ref2", 0)
                    objjournalentry.Reference3 = DateTime.Now.ToString();
                    objjournalentry.Memo = Convert.ToString(objRs.Fields.Item("Memo").Value) + "(Reversal) - " + JETransId; // objJEHeader.GetValue("Memo", 0) & " (Reversal) - " & Trim(JETransId)
                    objjournalentry.Series = Convert.ToInt32(objRs.Fields.Item("Series").Value); // objJEHeader.GetValue("Series", 0)
                    objjournalentry.UserFields.Fields.Item("U_RevRecDN").Value = Convert.ToString(objRs.Fields.Item("U_RevRecDN").Value);
                    objjournalentry.UserFields.Fields.Item("U_RevRecDE").Value = Convert.ToString(objRs.Fields.Item("U_RevRecDE").Value);
           
                for (int AccRow = 0; AccRow < objRs.RecordCount ; AccRow++)
                    {
                        if (Convert.ToString(objRs.Fields.Item("BPCode").Value) != "")
                            objjournalentry.Lines.ShortName = Convert.ToString(objRs.Fields.Item("BPCode").Value);
                        else
                            objjournalentry.Lines.AccountCode = Convert.ToString(objRs.Fields.Item("Account").Value);
                        if (Convert.ToDouble(objRs.Fields.Item("Credit").Value) != 0)
                            objjournalentry.Lines.Debit = Convert.ToDouble(objRs.Fields.Item("Credit").Value);
                        else
                            objjournalentry.Lines.Credit = Convert.ToDouble(objRs.Fields.Item("Debit").Value);
                        if(Convert.ToString(objRs.Fields.Item("BPLId").Value)!="") objjournalentry.Lines.BPLID = Convert.ToInt32(objRs.Fields.Item("BPLId").Value);
                        objjournalentry.Lines.UserFields.Fields.Item("U_InvEntry").Value =Convert.ToString( objRs.Fields.Item("U_InvEntry").Value); // Branch
                        objjournalentry.Lines.Add();
                        objRs.MoveNext();
                    }

                    if (objjournalentry.Add() != 0)
                    {
                        if (clsModule.objaddon.objcompany.InTransaction) clsModule.objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                        clsModule.objaddon.objapplication.SetStatusBarMessage("Journal Reverse: " + clsModule.objaddon.objcompany.GetLastErrorDescription() + "-" + clsModule.objaddon.objcompany.GetLastErrorCode(), SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(objjournalentry);
                        return false;
                    }
                    // 
                    else
                    {
                      if (clsModule.objaddon.objcompany.InTransaction) clsModule.objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                      TransId = clsModule.objaddon.objcompany.GetNewObjectKey();                
                        
                     ((SAPbouiCOM.EditText)objform.Items.Item("trvtran").Specific).String = TransId;
                     objRs.DoQuery("Update OJDT set \"StornoToTr\"=" + JETransId + " where \"TransId\"=" + TransId + "");
                    ((SAPbouiCOM.ComboBox)objform.Items.Item("cstatus").Specific).Select("C", SAPbouiCOM.BoSearchKey.psk_ByValue);
                    objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                    objform.Items.Item("1").Click();
                    objform.Items.Item("trvtran").Visible = true;
                    objform.Items.Item("lrvtran").Visible = true;
                    objform.Items.Item("lkrvtran").Visible = true;
                    objmatrix.Item.Enabled = false;
                    clsModule.objaddon.objapplication.StatusBar.SetText("Journal Entry Reversed Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                    return true;
                    }

                    //if (ErrorFlag)
                    //{
                    //    ((SAPbouiCOM.EditText)objform.Items.Item("trvtran").Specific).String = "";
                    //}
                    //else
                    //{
                    //    
                    //    clsModule.objaddon.objapplication.StatusBar.SetText("Transactions Cancelled Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                    //    return true;
                    //}
                }
                catch (Exception ex)
                {
                    if (clsModule.objaddon.objcompany.InTransaction)  clsModule.objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                    clsModule.objaddon.objapplication.SetStatusBarMessage("Transaction Cancelling Error " + clsModule.objaddon.objcompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                    return false;
                }
            

        }


    }
}
