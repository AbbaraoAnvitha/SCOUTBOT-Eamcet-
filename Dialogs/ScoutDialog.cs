using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Threading.Tasks;
using System.Data.OleDb;

using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Builder.FormFlow;
using Microsoft.Bot.Builder.Luis;
using Microsoft.Bot.Builder.Luis.Models;
using Microsoft.Bot.Connector;
using System.Data;
using System.Text;

namespace Scout.Dialogs
{
    [LuisModel("9f2949ee-4120-4433-a0be-0e633f339718", "c59ef1ca4a0e41e7ad228294d92ae597")]
    [Serializable]
    public class ScoutDialog : LuisDialog<object>
    {
        [LuisIntent("getIntro")]
        public async Task getIntro(IDialogContext context, LuisResult result)
        {
            string res = "Hi, I am scout and I will help you in finding colleges for your eamcet rank. <br />You can also get the colleges rank information by entering the college name <br />Example: Get the ranks of bvrith<br />";
            await context.PostAsync(res);
            context.Wait(MessageReceived);
        }

        [LuisIntent("getConclu")]
        public async Task getConclu(IDialogContext context, LuisResult result)
        {
            string res = "See you later.Hope I helped you :)";
            await context.PostAsync(res);
            context.Wait(MessageReceived);
        }

        [LuisIntent("getCollege")]
        public async Task getCollege(IDialogContext context, LuisResult result)
        {
            var rankEntity = result.Entities.SingleOrDefault(e => e.Type == "Rank");

            System.Data.DataTable dtExcel;
            dtExcel = new System.Data.DataTable();
            dtExcel.TableName = "MyExcelData";
            string SourceConstr = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source='~\eamcetdataset.xlsx';Extended Properties= 'Excel 8.0;HDR=Yes;IMEX=1'";
            OleDbConnection con = new OleDbConnection(SourceConstr);
            string query = "Select * from [Sheet1$]";
            OleDbDataAdapter data = new OleDbDataAdapter(query, con);
            data.Fill(dtExcel);

            string expression = "ocg >= " + Int32.Parse(rankEntity.Entity) + " AND ocg <="+Int32.Parse(rankEntity.Entity)+5000;
            DataRow[] foundRows;

            // Use the Select method to find all rows matching the filter.
            foundRows = dtExcel.Select(expression,"ocg ASC");

            //Print column 0 of each returned row.

            int k = 0;
            string res = "";
            for (int i = 0; i < foundRows.Length; i++)
            {
                k++;
                res += foundRows[i][0] +  "<br />";
                if(k == 20)
                {
                    break;
                }
            }


            await context.PostAsync(res);
            context.Wait(MessageReceived);
            
        }

        [LuisIntent("getBestBranch")]
        public async Task getBestBranch(IDialogContext context, LuisResult result)
        {
            var collegeEntity = result.Entities.SingleOrDefault(e => e.Type == "College");
            System.Data.DataTable dtExcel;
            dtExcel = new System.Data.DataTable();
            dtExcel.TableName = "MyExcelData";
            string SourceConstr = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source='~\eamcetdataset.xlsx';Extended Properties= 'Excel 8.0;HDR=Yes;IMEX=1'";
            OleDbConnection con = new OleDbConnection(SourceConstr);
            string query = "Select * from [Sheet1$]";
            OleDbDataAdapter data = new OleDbDataAdapter(query, con);
            data.Fill(dtExcel);

            string expression = "College = '" + collegeEntity.Entity.ToUpper() + "'";
            DataRow[] foundRows;

            // Use the Select method to find all rows matching the filter.
            foundRows = dtExcel.Select(expression,"ocg ASC");

            //Print column 0 of each returned row.
            string res = "";
                res += foundRows[0][1] + " " + Environment.NewLine;
          



            await context.PostAsync(res);
            context.Wait(MessageReceived);
        }

        [LuisIntent("getRank")]
        public async Task getRank(IDialogContext context, LuisResult result)
        {
          
            var collegeEntity = result.Entities.SingleOrDefault(e => e.Type == "College");

            System.Data.DataTable dtExcel;
            dtExcel = new System.Data.DataTable();
            dtExcel.TableName = "MyExcelData";
            string SourceConstr = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source='~\eamcetdataset.xlsx';Extended Properties= 'Excel 8.0;HDR=Yes;IMEX=1'";
            OleDbConnection con = new OleDbConnection(SourceConstr);
            string query = "Select * from [Sheet1$]";
            OleDbDataAdapter data = new OleDbDataAdapter(query, con);
            data.Fill(dtExcel);

            string expression = "College = '" + collegeEntity.Entity.ToUpper() + "'";
            DataRow[] foundRows;

            //Use the Select method to find all rows matching the filter.
            foundRows = dtExcel.Select(expression);

            //Print column 0 of each returned row.
            string res = "Branch\t\tOC\t\tSC\t\tST\t\tBC <br />";

            for (int i = 0; i < foundRows.Length; i++)
            {
                res += foundRows[i][1] + new string(' ',10) + foundRows[i][3] + new string(' ', 10) + foundRows[i][5] + new string(' ', 10) + foundRows[i][7] + new string(' ', 10) + foundRows[i][9] + "<br />";
            }


            await context.PostAsync(res);
            context.Wait(MessageReceived);
        }

        [LuisIntent("getBranches")]
        public async Task getBranches(IDialogContext context, LuisResult result)
        {

            var collegeEntity = result.Entities.SingleOrDefault(e => e.Type == "College");

            System.Data.DataTable dtExcel;
            dtExcel = new System.Data.DataTable();
                dtExcel.TableName = "MyExcelData";
                string SourceConstr = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source='~\eamcetdataset.xlsx';Extended Properties= 'Excel 8.0;HDR=Yes;IMEX=1'";
                OleDbConnection con = new OleDbConnection(SourceConstr);
            string query = "Select * from [Sheet1$]";
                OleDbDataAdapter data = new OleDbDataAdapter(query, con);
                data.Fill(dtExcel);
           
            string  expression = "College = '"+collegeEntity.Entity.ToUpper()+"'";
            DataRow[] foundRows;

            // Use the Select method to find all rows matching the filter.
            foundRows = dtExcel.Select(expression);

            //Print column 0 of each returned row.
            string res = "";
           
            for (int i = 0; i < foundRows.Length; i++)
            {
                res += foundRows[i][1] + "     "+ "<br />";
            }
            
            await context.PostAsync(res);
            context.Wait(MessageReceived);
        }
        [LuisIntent("getInformation")]
        public async Task getInformation(IDialogContext context, LuisResult result)
        {

            var infoEntity = result.Entities.SingleOrDefault(e => e.Type == "College");

            System.Data.DataTable dtExcel;
            dtExcel = new System.Data.DataTable();
            dtExcel.TableName = "MyExcelData";
            string SourceConstr = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source='~\eamcetdataset.xlsx';Extended Properties= 'Excel 8.0;HDR=Yes;IMEX=1'";
            OleDbConnection con = new OleDbConnection(SourceConstr);
            string query = "Select * from [Sheet1$]";
            OleDbDataAdapter data = new OleDbDataAdapter(query, con);
            data.Fill(dtExcel);

            string expression = "College = '" + infoEntity.Entity.ToUpper() + "'";
            DataRow[] foundRows;

            // Use the Select method to find all rows matching the filter.
            foundRows = dtExcel.Select(expression);

            //Print column 0 of each returned row.
            string res = "";


            res = "To get further information about the college you can visit this website link" + " " + foundRows[0][19] + "     " + "<br />";


            await context.PostAsync(res);
            context.Wait(MessageReceived);
        }

        [LuisIntent("getWomen")]
        public async Task getWomen(IDialogContext context, LuisResult result)
        {
            var womenEntity = result.Entities.SingleOrDefault(e => e.Type == "Womens");
            DataTable dtExcel = new DataTable();
            dtExcel.Clear();

            dtExcel.Columns.Add("College");
            dtExcel.Columns.Add("Region");

            dtExcel.Columns.Add("Rank");
            dtExcel.Columns.Add("Location");
            dtExcel.Columns.Add("Women");
            DataRow _r1 = dtExcel.NewRow();

            _r1["College"] = "JOGINPALLY MN RAO WOMENS ENGINEERING COLLEGE";
            _r1["Region"] = "OU";
            _r1["Rank"] = 77827;
            _r1["Location"] = "YENKAPALLY";
            _r1["Women"] = "women";
            dtExcel.Rows.Add(_r1);
            DataRow _r2 = dtExcel.NewRow();

            _r2["College"] = "G NARAYNAMMA INSTITUTE OF TECHNOLOGY AND SCIENCE";
            _r2["Region"] = "OU";
            _r2["Rank"] = 9756;
            _r2["Location"] = "RAYADURG";
            _r2["Women"] = "women";
            dtExcel.Rows.Add(_r2);
            DataRow _r3 = dtExcel.NewRow();

            _r3["College"] = "BHOJREDDY ENGINERING COLLEGE FOR WOMEN";
            _r3["Region"] = "OU";
            _r3["Location"] = "SAIDABAD";
            _r3["Rank"] = 21186;
            _r3["Women"] = "women";
            dtExcel.Rows.Add(_r3);
            DataRow _r4 = dtExcel.NewRow();
            _r4["College"] = "BVRIT COLLEGE OF ENGINEERING FOR WOMEN";
            _r4["Region"] = "OU";
            _r4["Location"] = "BACHUPALLY";
            _r4["Rank"] = 11556;
            _r4["Women"] = "women";
            dtExcel.Rows.Add(_r4);
            DataRow _r5 = dtExcel.NewRow();

            _r5["College"] = "MALLA REDDY WOMENS ENGINEERING COLLEGE";
            _r5["Region"] = "OU";
            _r5["Location"] = "MAISAMMAGUDA";
            _r5["Rank"] = 83450;
            _r5["Women"] = "women";
            DataRow _r6 = dtExcel.NewRow();

            _r6["College"] = "SAHASRA COLLEGE OF ENGINEERING FOR WOMEN";
            _r6["Region"] = "OU";
            _r6["Location"] = "WARANGAL";
            _r6["Rank"] = 25670;
            _r6["Women"] = "women";
            DataRow _r7 = dtExcel.NewRow();

            _r7["College"] = "SRIDEVI WOMENS ENGINEERING COLLEGE";
            _r7["Region"] = "OU";
            _r7["Location"] = "GANDIPET";
            _r7["Rank"] = 83569;
            _r7["Women"] = "women";
            DataRow _r8 = dtExcel.NewRow();

            _r8["College"] = "STANLEY COLLEGE OF ENGINEERING AND TECHNOLOGY FOR WOMEN";
            _r8["Region"] = "OU";
            _r8["Location"] = "ABIDS";
            _r8["Rank"] = 73438;
            _r8["Women"] = "women";
            DataRow _r9 = dtExcel.NewRow();

            _r9["College"] = "VIJAY COLLEGE OF ENGINEERING AND TECHNOLOGY FOR WOMEN";
            _r9["Region"] = "OU";
            _r9["Location"] = "NIZAMBAD";
            _r9["Rank"] = 206570;
            _r9["Women"] = "women";
            string expression = "Women = '" + womenEntity.Entity + "'";

            DataRow[] foundRows;

            // Use the Select method to find all rows matching the filter.
            foundRows = dtExcel.Select(expression);

            //Print column 0 of each returned row.
            string res = "";

            for (int i = 0; i < foundRows.Length; i++)
            {
                res += foundRows[i][0] + "    " + "<br />";
            }
            await context.PostAsync(res);
            context.Wait(MessageReceived);
        }



        [LuisIntent("")]
        [LuisIntent("None")]
        public async Task None(IDialogContext context, LuisResult result)
        {
            await context.PostAsync("Sorry I didn't understand that");
            context.Wait(MessageReceived);
        }
    }
}