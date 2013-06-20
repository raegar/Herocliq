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
    public partial class WebForm1 : System.Web.UI.Page
    {
        public TextBox[] speed = new TextBox[12];
        public TextBox[] attack = new TextBox[12];
        public TextBox[] defence = new TextBox[12];
        public TextBox[] damage = new TextBox[12];
        public TextBox txtCliqName = new TextBox();
        public TextBox txtCliqTeam = new TextBox();
        public TextBox txtCliqRange = new TextBox();
        public TextBox txtCliqPoints = new TextBox();
        public TextBox txtCliqKeywords = new TextBox();
        public System.Web.UI.WebControls.Image[] icons = new System.Web.UI.WebControls.Image[4];

        protected void Page_Load(object sender, EventArgs e)
        {
            //Create Event Handler
            EventHandler handler = new EventHandler(handleTextClicks_Click); 

            //Create panel with HTML table
            int counter = 12;
            Panel1.Controls.Add(new LiteralControl("<table width=50%><tr>"));
            Panel1.Controls.Add(new LiteralControl("<td class='style2' colspan = '12'>"));
            Panel1.Controls.Add(new LiteralControl("<h2>Name - "));
            Panel1.Controls.Add(txtCliqName);
            Panel1.Controls.Add(new LiteralControl("</h2><h2>Team - "));
            Panel1.Controls.Add(txtCliqTeam);
            Panel1.Controls.Add(new LiteralControl("</h2><h2>Range - "));
            Panel1.Controls.Add(txtCliqRange);
            Panel1.Controls.Add(new LiteralControl("</h2><h2>Points - "));
            Panel1.Controls.Add(txtCliqPoints);
            Panel1.Controls.Add(new LiteralControl("</h2><h2>Keywords - "));
            Panel1.Controls.Add(txtCliqKeywords);
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
                speed[i] = new TextBox();
                speed[i].Attributes.Add("style", "text-align: center;");
                speed[i].Width = 20;
                speed[i].Height = 20;
                speed[i].Font.Bold = true;
                speed[i].TextChanged += new EventHandler(handler); ////NEED TO WORK OUT HANDLER
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
                attack[i] = new TextBox();
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
                defence[i] = new TextBox();
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
                damage[i] = new TextBox();
                damage[i].Attributes.Add("style", "text-align: center;");
                damage[i].Width = 20;
                damage[i].Height = 20;
                damage[i].Font.Bold = true;
                Panel1.Controls.Add(new LiteralControl("<td>"));
                Panel1.Controls.Add(damage[i]);
                Panel1.Controls.Add(new LiteralControl("</td>"));
            }

            //Set colour picker stats
            Panel1.Controls.Add(new LiteralControl("</tr></table>"));
        }


        protected void Button1_Click(object sender, EventArgs e)
        {

        }

        protected void handleTextClicks_Click(object sender, EventArgs e)
        {
            TextBox tb = new TextBox();
            tb = (TextBox)sender;
            tb.BackColor = Color.Green;
        }
    }
}