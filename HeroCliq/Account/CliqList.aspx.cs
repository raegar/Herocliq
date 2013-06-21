using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data.SqlClient;
using System.Data;
using System.Drawing;

namespace HeroCliq.Account
{
    public partial class CliqList : System.Web.UI.Page
    {
        public Label[] speed = new Label[12];
        public Label[] attack = new Label[12];
        public Label[] defence = new Label[12];
        public Label[] damage = new Label[12];
        public Label lblCliqName = new Label();
        public Label lblCliqTeam = new Label();
        public Label lblCliqRange = new Label();
        public Label lblCliqPoints = new Label();
        public Label lblCliqKeywords = new Label();
        public System.Web.UI.WebControls.Image[] icons = new System.Web.UI.WebControls.Image[4];

        protected void Page_Load(object sender, EventArgs e)
        {
            //Create panel with HTML table
            int counter = 12;
            Panel1.Controls.Add(new LiteralControl("<table width=50%><tr>"));
            Panel1.Controls.Add(new LiteralControl("<td class='style2' colspan = '12'>"));
            Panel1.Controls.Add(new LiteralControl("<h2>Name - "));
            Panel1.Controls.Add(lblCliqName);
            Panel1.Controls.Add(new LiteralControl("</h2><h2>Team - "));
            Panel1.Controls.Add(lblCliqTeam);
            Panel1.Controls.Add(new LiteralControl("</h2><h2>Range - "));
            Panel1.Controls.Add(lblCliqRange);
            Panel1.Controls.Add(new LiteralControl("</h2><h2>Points - "));
            Panel1.Controls.Add(lblCliqPoints);
            Panel1.Controls.Add(new LiteralControl("</h2><h2>Keywords - "));
            Panel1.Controls.Add(lblCliqKeywords);
            Panel1.Controls.Add(new LiteralControl("<tr></tr>")); 

            //Add Speed icon
            icons[0] = new System.Web.UI.WebControls.Image();
            icons[0].ImageUrl = @"..\Images\Speed.png";
            Panel1.Controls.Add(new LiteralControl("<td>"));
            Panel1.Controls.Add(icons[0]);
            Panel1.Controls.Add(new LiteralControl("</td>"));

            //Add Strength stats
            for (int i = 0; i < counter; i++)
            {
                speed[i] = new Label();
                speed[i].Attributes.Add("style", "text-align: center;");
                speed[i].Width = 20;
                speed[i].Height = 20;
                speed[i].Font.Bold = true;
                Panel1.Controls.Add(new LiteralControl("<td>"));
                Panel1.Controls.Add(speed[i]);
                Panel1.Controls.Add(new LiteralControl("</td>"));
            }
            Panel1.Controls.Add(new LiteralControl("</tr><tr>"));

            //Add Attack icon
            icons[1] = new System.Web.UI.WebControls.Image();
            icons[1].ImageUrl = @"..\Images\Attack.png";
            Panel1.Controls.Add(new LiteralControl("<td>"));
            Panel1.Controls.Add(icons[1]);
            Panel1.Controls.Add(new LiteralControl("</td>"));

            //Add Attack stats
            for (int i = 0; i < counter; i++)
            {
                attack[i] = new Label();
                attack[i].Attributes.Add("style", "text-align: center;");
                attack[i].Width = 20;
                attack[i].Height = 20;
                attack[i].Font.Bold = true;
                Panel1.Controls.Add(new LiteralControl("<td>"));
                Panel1.Controls.Add(attack[i]);
                Panel1.Controls.Add(new LiteralControl("</td>"));
            }
            Panel1.Controls.Add(new LiteralControl("</tr><tr>"));

            //Add Defence icon
            icons[2] = new System.Web.UI.WebControls.Image();
            icons[2].ImageUrl = @"..\Images\Defence.png";
            Panel1.Controls.Add(new LiteralControl("<td>"));
            Panel1.Controls.Add(icons[2]);
            Panel1.Controls.Add(new LiteralControl("</td>"));

            //Add Defence stats
            for (int i = 0; i < counter; i++)
            {
                defence[i] = new Label();
                defence[i].Attributes.Add("style", "text-align: center;");
                defence[i].Width = 20;
                defence[i].Height = 20;
                defence[i].Font.Bold = true;
                Panel1.Controls.Add(new LiteralControl("<td>"));
                Panel1.Controls.Add(defence[i]);
                Panel1.Controls.Add(new LiteralControl("</td>"));
            }
            Panel1.Controls.Add(new LiteralControl("</tr><tr>"));

            //Add Damage icon
            icons[3] = new System.Web.UI.WebControls.Image();
            icons[3].ImageUrl = @"..\Images\Damage.png";
            Panel1.Controls.Add(new LiteralControl("<td>"));
            Panel1.Controls.Add(icons[3]);
            Panel1.Controls.Add(new LiteralControl("</td>"));


            //Add Damage stats
            for (int i = 0; i < counter; i++)
            {
                damage[i] = new Label();
                damage[i].Attributes.Add("style", "text-align: center;");
                damage[i].Width = 20;
                damage[i].Height = 20;
                damage[i].Font.Bold = true;
                Panel1.Controls.Add(new LiteralControl("<td>"));
                Panel1.Controls.Add(damage[i]);
                Panel1.Controls.Add(new LiteralControl("</td>"));
            }
            Panel1.Controls.Add(new LiteralControl("</tr></table>"));
        }

        protected void lstCliqs_SelectedIndexChanged(object sender, EventArgs e)
        {
           
            string cliqNumber = lstCliqs.Text;

            string myConnectionString = @"Data Source=.\SQLEXPRESS;AttachDbFilename=|DataDirectory|\HeroCliq_Cliqs.mdf;Integrated Security=True;Connect Timeout=30;User Instance=True;";
            SqlConnection conn = new SqlConnection(myConnectionString);
            SqlCommand comm = new SqlCommand("SELECT * FROM Cliqs WHERE Cliq_Number LIKE '" + cliqNumber + "'", conn);
            conn.Open();
            SqlDataReader reader = comm.ExecuteReader();
            while (reader.Read())
            {
                lblCliqName.Text = (string)reader["Cliq_Number"] + " " + reader["Cliq_Name"];
                lblCliqTeam.Text = (string)reader["Cliq_Team"];
                lblCliqRange.Text = "" + reader["Cliq_Range"];
                lblCliqPoints.Text = "" + reader["Cliq_Points"];
                lblCliqKeywords.Text = (string)reader["Cliq_Keyword"];
                for (int i = 0; i < 12; i++)
                {
                    speed[i].Text = splitData((string)reader["Cliq_Speed"])[i];
                    setColour(speed[i], i, splitData((string)reader["Cliq_Speed_Colour"])[i]);
                    attack[i].Text = splitData((string)reader["Cliq_Attack"])[i];
                    setColour(attack[i], i, splitData((string)reader["Cliq_Attack_Colour"])[i]);
                    defence[i].Text = splitData((string)reader["Cliq_Defence"])[i];
                    setColour(defence[i], i, splitData((string)reader["Cliq_Defence_Colour"])[i]);
                    damage[i].Text = splitData((string)reader["Cliq_Damage"])[i];
                    setColour(damage[i], i, splitData((string)reader["Cliq_Damage_Colour"])[i]);
                }
                   
        
            }
            reader.Close();
            conn.Close();

            
            

        }

        public void setColour(Label cliqLabel, int i, string value)
        {

            switch (value)
            {

                case "Re":
                    cliqLabel.BackColor = Color.Red;
                    cliqLabel.ForeColor = Color.White;
                    break;
                case "Gr":
                    cliqLabel.BackColor = Color.Gray;
                    cliqLabel.ForeColor = Color.White;
                break;
                case "Pu":
                    cliqLabel.BackColor = Color.Purple;
                    cliqLabel.ForeColor = Color.White;
                break;
                case "Br":
                    cliqLabel.BackColor = Color.SaddleBrown;
                    cliqLabel.ForeColor = Color.White;
                break;
                case "Or":
                    cliqLabel.BackColor = Color.DarkOrange;
                    cliqLabel.ForeColor = Color.Black;
                break;
                case "Bl":
                    cliqLabel.BackColor = Color.Black;
                    cliqLabel.ForeColor = Color.White;
                break;
                case "LB":
                    cliqLabel.BackColor = Color.LightBlue;
                    cliqLabel.ForeColor = Color.Black;
                break;
                case "DB":
                    cliqLabel.BackColor = Color.DarkBlue;
                    cliqLabel.ForeColor = Color.White;
                break;
                case "LG":
                    cliqLabel.BackColor = Color.OliveDrab;
                    cliqLabel.ForeColor = Color.Black;
                break;
                case "DG":
                    cliqLabel.BackColor = Color.Green;
                    cliqLabel.ForeColor = Color.White;
                break;
                case "Wh":
                    cliqLabel.BackColor = Color.White;
                    cliqLabel.ForeColor = Color.Black;
                    cliqLabel.BorderStyle = BorderStyle.Solid;
                    cliqLabel.BorderWidth = 2;
                break;
                case "KO":
                cliqLabel.BackColor = Color.White;
                cliqLabel.ForeColor = Color.Red;
                break;
                default:
                    cliqLabel.BackColor = Color.White;
                    cliqLabel.ForeColor = Color.Black;
                break;

                    

            }
        }

        public List<string> splitData(string data)
        {
            string[] dataArray = data.Split(',');
            List<string> dataList = new List<string>(dataArray.Length);
            dataList.AddRange(dataArray);

            return dataList;
        }

    }
}