# FanduelAlgorithm
Generates Rosters that are optimized for projected points while still within salary cap.


using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
//using Microsoft.Office.Interop.Excel;
using Microsoft.Office;
using System.Reflection;
using System.Data.OleDb;

namespace FanDuel_Algorithm
{
    public partial class Form1 : Form
    {
        //SET THE SALARY CAP
        int SalaryCap = 60000;



        public Form1()
        {
            InitializeComponent();


        }

        private void Form1_Load(object sender, EventArgs e)
        {


        }

//this is the holy grail button, MAKE SURE THIS WORKS!!!
        private void button1_Click(object sender, EventArgs e)
        {
            //populate the array of quarterbacks
            string[,] QBArray = new string[dataGridView1.Rows.Count, dataGridView1.Columns.Count];

            for (int x = 0; x < QBArray.GetLength(0); x++)
            {
                for (int i = 0; i < QBArray.GetLength(1); i++)
                {
                    QBArray[x, i] = Convert.ToString(dataGridView1.Rows[x].Cells[i].Value);
                }
            }

            //populate the array of runningbacks
            string[,] RBArray = new string[dataGridView2.Rows.Count, dataGridView2.Columns.Count];

            for (int x = 0; x < RBArray.GetLength(0); x++)
            {
                for (int i = 0; i < RBArray.GetLength(1); i++)
                {
                    RBArray[x, i] = Convert.ToString(dataGridView2.Rows[x].Cells[i].Value);
                }
            }

            //populate the array of wide receivers
            string[,] WRArray = new string[dataGridView3.Rows.Count, dataGridView3.Columns.Count];

            for (int x = 0; x < WRArray.GetLength(0); x++)
            {
                for (int i = 0; i < WRArray.GetLength(1); i++)
                {
                    WRArray[x, i] = Convert.ToString(dataGridView3.Rows[x].Cells[i].Value);
                }
            }

            //populate the array of tight ends
            string[,] TEArray = new string[dataGridView4.Rows.Count, dataGridView4.Columns.Count];

            for (int x = 0; x < TEArray.GetLength(0); x++)
            {
                for (int i = 0; i < TEArray.GetLength(1); i++)
                {
                    TEArray[x, i] = Convert.ToString(dataGridView4.Rows[x].Cells[i].Value);
                }
            }

            //populate the array of defenses
            string[,] DEFArray = new string[dataGridView5.Rows.Count, dataGridView5.Columns.Count];

            for (int x = 0; x < DEFArray.GetLength(0); x++)
            {
                for (int i = 0; i < DEFArray.GetLength(1); i++)
                {
                    DEFArray[x, i] = Convert.ToString(dataGridView5.Rows[x].Cells[i].Value);
                }
            }

            //populate the array of kickers
            string[,] KArray = new string[dataGridView6.Rows.Count, dataGridView6.Columns.Count];

            for (int x = 0; x < KArray.GetLength(0); x++)
            {
                for (int i = 0; i < KArray.GetLength(1); i++)
                {
                    KArray[x, i] = Convert.ToString(dataGridView6.Rows[x].Cells[i].Value);
                }
            }

            //find how many players are available at each position
            //set those variables
            int QBtot = QBArray.GetLength(0);
            int RBtot = RBArray.GetLength(0);
            int WRtot = WRArray.GetLength(0);
            int TEtot = TEArray.GetLength(0);
            int DEFtot = DEFArray.GetLength(0);
            int Ktot = KArray.GetLength(0);



            //establish roster array

            string[, ,] RosterArray = new string[11, 10, 3];
            for (int i = 0; i < 11; i++)
            {
                for (int j = 0; j < 10; j++)
                {
                    for (int k = 0; k < 3; k++)
                    {
                        RosterArray[i, j, k] = "0";
                    }
                }
            }


            Int64 totalsalary = 0;
            double totalpoints;

            for (int a = 0; a < QBtot; a++)
            {
                for (int b = 0; b < RBtot; b++)
                {
                    for (int c = (b+1); c < RBtot; c++)
                    {
                            for (int d = 0; d < WRtot; d++)
                            {
                                for (int f = (d+1); f < WRtot; f++)
                                {
                                        for (int g = (f+1); g < WRtot; g++)
                                        {
                                                for (int h = 0; h < TEtot; h++)
                                                {
                                                    for (int i = 0; i < DEFtot; i++)
                                                    {
                                                        for (int j = 0; j < Ktot; j++)
                                                        {
                                                          //make row 0 the comparison row
                                                          
                                                          //populate quarterbacks
                                                            RosterArray[0, 0, 0] = QBArray[a, 0];
                                                            RosterArray[0, 0, 1] = QBArray[a, 1];
                                                            RosterArray[0, 0, 2] = QBArray[a, 2];
                                                          
                                                          //populate runningback1
                                                            RosterArray[0, 1, 0] = RBArray[b, 0];
                                                            RosterArray[0, 1, 1] = RBArray[b, 1];
                                                            RosterArray[0, 1, 2] = RBArray[b, 2];
                                                          
                                                          //populate runningback2
                                                            RosterArray[0, 2, 0] = RBArray[c, 0];
                                                            RosterArray[0, 2, 1] = RBArray[c, 1];
                                                            RosterArray[0, 2, 2] = RBArray[c, 2];

                                                          //populate wide receiver1
                                                            RosterArray[0, 3, 0] = WRArray[d, 0];
                                                            RosterArray[0, 3, 1] = WRArray[d, 1];
                                                            RosterArray[0, 3, 2] = WRArray[d, 2];

                                                          //populate wide receiver2
                                                            RosterArray[0, 4, 0] = WRArray[f, 0];
                                                            RosterArray[0, 4, 1] = WRArray[f, 1];
                                                            RosterArray[0, 4, 2] = WRArray[f, 2];
                                                            
                                                          //populate wide receiver3
                                                            RosterArray[0, 5, 0] = WRArray[g, 0];
                                                            RosterArray[0, 5, 1] = WRArray[g, 1];
                                                            RosterArray[0, 5, 2] = WRArray[g, 2];

                                                          //populate tight ends
                                                            RosterArray[0, 6, 0] = TEArray[h, 0];
                                                            RosterArray[0, 6, 1] = TEArray[h, 1];
                                                            RosterArray[0, 6, 2] = TEArray[h, 2];

                                                          //populate defenses
                                                            RosterArray[0, 7, 0] = DEFArray[i, 0];
                                                            RosterArray[0, 7, 1] = DEFArray[i, 1];
                                                            RosterArray[0, 7, 2] = DEFArray[i, 2];

                                                          //populate kickers
                                                            RosterArray[0, 8, 0] = KArray[j, 0];
                                                            RosterArray[0, 8, 1] = KArray[j, 1];
                                                            RosterArray[0, 8, 2] = KArray[j, 2];
                                                            
    
            //calculate the salary of the team
totalsalary = Convert.ToInt64(QBArray[a, 1]) + Convert.ToInt64(RBArray[b, 1]) + Convert.ToInt64(RBArray[c, 1]) +
             Convert.ToInt64(WRArray[d, 1]) + Convert.ToInt64(WRArray[f, 1]) + Convert.ToInt64(WRArray[g, 1]) +
             Convert.ToInt64(TEArray[h, 1]) + Convert.ToInt64(DEFArray[i, 1]) + Convert.ToInt64(KArray[j, 1]);

            //calculate the point projections of the team
totalpoints = Convert.ToDouble(QBArray[a, 2]) + Convert.ToDouble(RBArray[b, 2]) + Convert.ToDouble(RBArray[c, 2]) +
              Convert.ToDouble(WRArray[d, 2]) + Convert.ToDouble(WRArray[f, 2]) + Convert.ToDouble(WRArray[g, 2]) +
              Convert.ToDouble(TEArray[h, 2]) + Convert.ToDouble(DEFArray[i, 2]) + Convert.ToDouble(KArray[j, 2]);
              
              //set the values of the total block
              RosterArray[0, 9, 0] = "Total: ";
              RosterArray[0, 9, 1] = Convert.ToString(totalsalary);
              RosterArray[0, 9, 2] = Convert.ToString(totalpoints);


                  //check if the team salary is less than the salary cap
                  if (totalsalary <= SalaryCap)
                    {
                    // increment loop to check the existing top 10 lineups you already have
                      for (int z = 1; z < 11; z++)
                          {
                          //check if the points of the team in the comparison block is greater
                          //than the points in the z block
                            if (totalpoints > Convert.ToDouble(RosterArray[z, 9, 2]))
                                {
                                //check to see if I can proceed
                                  if (z != 11)
                                    {
                                    //if I can proceed, start from the back end of my top 10 teams
                                    //and shift them all down one block in the array
                                      for (int y = 9; y >= z; y--)
                                         {
                                         //shift the quarterbacks down one block
                                          RosterArray[(y + 1), 0, 0] = RosterArray[y, 0, 0];
                                          RosterArray[(y + 1), 0, 1] = RosterArray[y, 0, 1];
                                          RosterArray[(y + 1), 0, 2] = RosterArray[y, 0, 2];

                                          //shift runningback1 down one block
                                          RosterArray[(y + 1), 1, 0] = RosterArray[y, 1, 0];
                                          RosterArray[(y + 1), 1, 1] = RosterArray[y, 1, 1];
                                          RosterArray[(y + 1), 1, 2] = RosterArray[y, 1, 2];

                                          //shift runningback2 down one block
                                          RosterArray[(y + 1), 2, 0] = RosterArray[y, 2, 0];
                                          RosterArray[(y + 1), 2, 1] = RosterArray[y, 2, 1];
                                          RosterArray[(y + 1), 2, 2] = RosterArray[y, 2, 2];

                                          //shift wide receiver1 down one block
                                          RosterArray[(y + 1), 3, 0] = RosterArray[y, 3, 0];
                                          RosterArray[(y + 1), 3, 1] = RosterArray[y, 3, 1];
                                          RosterArray[(y + 1), 3, 2] = RosterArray[y, 3, 2];
                                          
                                          //shift wide receiver2 down one block
                                          RosterArray[(y + 1), 4, 0] = RosterArray[y, 4, 0];
                                          RosterArray[(y + 1), 4, 1] = RosterArray[y, 4, 1];
                                          RosterArray[(y + 1), 4, 2] = RosterArray[y, 4, 2];

                                          //shift wide receiver3 down one block
                                           RosterArray[(y + 1), 5, 0] = RosterArray[y, 5, 0];
                                           RosterArray[(y + 1), 5, 1] = RosterArray[y, 5, 1];
                                           RosterArray[(y + 1), 5, 2] = RosterArray[y, 5, 2];

                                          //shift tight ends down one block
                                          RosterArray[(y + 1), 6, 0] = RosterArray[y, 6, 0];
                                          RosterArray[(y + 1), 6, 1] = RosterArray[y, 6, 1];
                                          RosterArray[(y + 1), 6, 2] = RosterArray[y, 6, 2];
                                          
                                          //shift defenses down one block
                                          RosterArray[(y + 1), 7, 0] = RosterArray[y, 7, 0];
                                          RosterArray[(y + 1), 7, 1] = RosterArray[y, 7, 1];
                                          RosterArray[(y + 1), 7, 2] = RosterArray[y, 7, 2];

                                          //shift kickers down one block
                                          RosterArray[(y + 1), 8, 0] = RosterArray[y, 8, 0];
                                          RosterArray[(y + 1), 8, 1] = RosterArray[y, 8, 1];
                                          RosterArray[(y + 1), 8, 2] = RosterArray[y, 8, 2];
                                          
                                          //shift totals down one block
                                          RosterArray[(y + 1), 9, 0] = RosterArray[y, 9, 0];
                                          RosterArray[(y + 1), 9, 1] = RosterArray[y, 9, 1];
                                          RosterArray[(y + 1), 9, 2] = RosterArray[y, 9, 2];
                                    
                            }

                                    //make the same changes as above by inserting the comparison 0 block
                                    //into the z block

                                     RosterArray[z, 0, 0] = RosterArray[0, 0, 0];
                                     RosterArray[z, 0, 1] = RosterArray[0, 0, 1];
                                     RosterArray[z, 0, 2] = RosterArray[0, 0, 2];

                                     RosterArray[z, 1, 0] = RosterArray[0, 1, 0];
                                     RosterArray[z, 1, 1] = RosterArray[0, 1, 1];
                                     RosterArray[z, 1, 2] = RosterArray[0, 1, 2];

                                     RosterArray[z, 2, 0] = RosterArray[0, 2, 0];
                                     RosterArray[z, 2, 1] = RosterArray[0, 2, 1];
                                     RosterArray[z, 2, 2] = RosterArray[0, 2, 2];

                                     RosterArray[z, 3, 0] = RosterArray[0, 3, 0];
                                     RosterArray[z, 3, 1] = RosterArray[0, 3, 1];
                                     RosterArray[z, 3, 2] = RosterArray[0, 3, 2];

                                     RosterArray[z, 4, 0] = RosterArray[0, 4, 0];
                                     RosterArray[z, 4, 1] = RosterArray[0, 4, 1];
                                     RosterArray[z, 4, 2] = RosterArray[0, 4, 2];

                                     RosterArray[z, 5, 0] = RosterArray[0, 5, 0];
                                     RosterArray[z, 5, 1] = RosterArray[0, 5, 1];
                                     RosterArray[z, 5, 2] = RosterArray[0, 5, 2];

                                     RosterArray[z, 6, 0] = RosterArray[0, 6, 0];
                                     RosterArray[z, 6, 1] = RosterArray[0, 6, 1];
                                     RosterArray[z, 6, 2] = RosterArray[0, 6, 2];

                                     RosterArray[z, 7, 0] = RosterArray[0, 7, 0];
                                     RosterArray[z, 7, 1] = RosterArray[0, 7, 1];
                                     RosterArray[z, 7, 2] = RosterArray[0, 7, 2];

                                     RosterArray[z, 8, 0] = RosterArray[0, 8, 0];
                                     RosterArray[z, 8, 1] = RosterArray[0, 8, 1];
                                     RosterArray[z, 8, 2] = RosterArray[0, 8, 2];

                                     RosterArray[z, 9, 0] = RosterArray[0, 9, 0];
                                     RosterArray[z, 9, 1] = RosterArray[0, 9, 1];
                                     RosterArray[z, 9, 2] = RosterArray[0, 9, 2];
                               }
                
                
                            //after replacing the z block with the comparison 0 block, set z to 11
                            //so double counting does not occur on roster spots, and the loop can now exit
                           z = 11;
                          }
                        }
                                                                
                     }
                    }
                  }
                }
              }
            }
          }
        }
      }
    }
            



//export to excel... finally
            
            //create 2D array to export to excel sheet
            string[,] WinningTeams = new string[11, 20];

            for (int q = 0; q < 9; q++)
            {
                for (int p = 1; p < 11; p++)
                {
                //set 2d winningteams array equal to relevant components in 3d rosterarray
                    WinningTeams[q, (p-1)] = RosterArray[p, q, 0];
                    WinningTeams[9, (p-1)] = RosterArray[p, 9, 1];
                    WinningTeams[10, (p-1)] = RosterArray[p, 9, 2];
                }
                
            
            
            
            }


            //open the excel application
            
            Excel.Application xlApp = new Excel.Application();
            xlApp.Visible = true;
            Excel.Workbook wb = xlApp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
            Excel.Worksheet ws = (Excel.Worksheet)wb.Worksheets.get_Item(1);
            Excel.Range rng = ws.Cells.get_Resize(WinningTeams.GetLength(0), WinningTeams.GetLength(1));
            rng.Value2 = WinningTeams;


        }

//button on the form that selects the excel sheet to import the player data from
        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                this.textBox_path.Text = openFileDialog1.FileName;
            }
        }

  //import the excel sheets for players into datagridviews that we can manipulate in arrays
  
        private void button3_Click(object sender, EventArgs e)
        {

            string PathConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + textBox_path.Text + ";Extended Properties=\"Excel 8.0;HDR=Yes;\";";
            OleDbConnection conn = new OleDbConnection(PathConn);

            OleDbDataAdapter myDataAdapter = new OleDbDataAdapter("Select * from [" + textBox_sheet.Text + "$]", conn);
            DataTable dt = new DataTable();

            myDataAdapter.Fill(dt);
            dataGridView1.DataSource = dt;


            string PathConn2 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + textBox_path.Text + ";Extended Properties=\"Excel 8.0;HDR=Yes;\";";
            OleDbConnection conn2 = new OleDbConnection(PathConn2);

            OleDbDataAdapter myDataAdapter2 = new OleDbDataAdapter("Select * from [" + textBox_sheet1.Text + "$]", conn2);
            DataTable dt2 = new DataTable();

            myDataAdapter2.Fill(dt2);
            dataGridView2.DataSource = dt2;



            string PathConn3 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + textBox_path.Text + ";Extended Properties=\"Excel 8.0;HDR=Yes;\";";
            OleDbConnection conn3 = new OleDbConnection(PathConn3);

            OleDbDataAdapter myDataAdapter3 = new OleDbDataAdapter("Select * from [" + textBox_sheet2.Text + "$]", conn3);
            DataTable dt3 = new DataTable();

            myDataAdapter3.Fill(dt3);
            dataGridView3.DataSource = dt3;


            string PathConn4 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + textBox_path.Text + ";Extended Properties=\"Excel 8.0;HDR=Yes;\";";
            OleDbConnection conn4 = new OleDbConnection(PathConn4);

            OleDbDataAdapter myDataAdapter4 = new OleDbDataAdapter("Select * from [" + textBox_sheet3.Text + "$]", conn4);
            DataTable dt4 = new DataTable();

            myDataAdapter4.Fill(dt4);
            dataGridView4.DataSource = dt4;

          
            string PathConn5 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + textBox_path.Text + ";Extended Properties=\"Excel 8.0;HDR=Yes;\";";
            OleDbConnection conn5 = new OleDbConnection(PathConn5);

            OleDbDataAdapter myDataAdapter5 = new OleDbDataAdapter("Select * from [" + textBox_sheet4.Text + "$]", conn5);
            DataTable dt5 = new DataTable();

            myDataAdapter5.Fill(dt5);
            dataGridView5.DataSource = dt5;


            string PathConn6 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + textBox_path.Text + ";Extended Properties=\"Excel 8.0;HDR=Yes;\";";
            OleDbConnection conn6 = new OleDbConnection(PathConn6);

            OleDbDataAdapter myDataAdapter6 = new OleDbDataAdapter("Select * from [" + textBox_sheet5.Text + "$]", conn6);
            DataTable dt6 = new DataTable();

            myDataAdapter6.Fill(dt6);
            dataGridView6.DataSource = dt6;

        }
  }
}
