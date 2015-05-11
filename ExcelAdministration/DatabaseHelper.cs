﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MemberAdministration
{
    class DatabaseHelper
    {
        private String dbPath;

        /// <summary>
        /// Setting the Database Path and saving it to the Properties of this Application. It will be loaded on the next Application start automatically.
        /// </summary>
        /// <param name="dbPath">Path to the Database which will be used for the connection.</param>
        public void setDatabase(String dbPath)
        {
            this.dbPath = dbPath;
            MemberAdministration.Properties.Settings.Default.dbPath = dbPath;
            Properties.Settings.Default.Save();
        }

        /// <summary>
        /// Pulling all members from the Excel file and returning a DataSet which will be used as a DataSource in the GridView
        /// </summary>
        /// <returns> DataSet </returns>
        public DataSet getAllMembers()
        {
            dbPath = MemberAdministration.Properties.Settings.Default.dbPath;
            if (dbPath != null)
            {
                /*
                 * Connection begins here when dbPath is set   ********   *********   *********
                 */
                String connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + dbPath + "; Extended Properties='Excel 12.0 Xml;HDR=YES'";

                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    try
                    {
                        connection.Open();

                        //this next line assumes that the file is in default Excel format with Sheet1 as the first sheet name, we need to adjust accordingly
                        OleDbDataAdapter adapter = new OleDbDataAdapter("SELECT * FROM [Sheet1$]", connection);
                        DataSet ds = new DataSet();
                        adapter.Fill(ds);
                        return ds;
                    }
                    catch (OleDbException)
                    {
                        Console.WriteLine("Something went wrong with DB connecting");
                    }
                }
            }
            return null;
        }

        /*
         * We need this for getting the specified user in the input box, either by name or surname 
         */
        public DataSet getThisMember(String member)
        {

            dbPath = MemberAdministration.Properties.Settings.Default.dbPath;
            if (dbPath != null)
            {
                /*
                 * Connection begins here when dbPath is set   ********   *********   *********
                 */
                String connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + dbPath + "; Extended Properties='Excel 12.0 Xml;HDR=YES'";

                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    try
                    {
                        connection.Open();
                        //this next line assumes that the file is in default Excel format with Sheet1 as the first sheet name, adjust accordingly
                        OleDbDataAdapter adapter = new OleDbDataAdapter("SELECT * FROM [Sheet1$] WHERE Nachname LIKE @member OR Vorname LIKE @member OR Adresse LIKE @member OR Ort LIKE @member", connection);
                        adapter.SelectCommand.Parameters.Add("@member", OleDbType.VarChar, 80).Value = "%" + member + "%";
                        DataSet ds = new DataSet();
                        adapter.Fill(ds);
                        return ds;

                    }
                    catch (OleDbException)
                    {
                        Console.WriteLine("Something went wrong with DB connecting for single member");
                    }
                }
            }
            return null;
        }

        /*
         *  We need to secure the connection first so that the INSERT functions properly. 
        */
        public void insertNewMember(Person person)
        {
            String ID = person.ID;
            String title = person.Title;
            String surname = person.Surname;
            String name = person.Name;
            String address = person.Address;
            String plz = person.Plz;
            String state = person.State;
            Int32 telephone = person.Telephone;
            String mail = person.Mail;
            String amount = person.Amount;

            dbPath = MemberAdministration.Properties.Settings.Default.dbPath;
            if (dbPath != null)
            {
                /*
                 * Connection begins here when dbPath is set   ********   *********   *********
                 */
                String connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + dbPath + "; Extended Properties='Excel 12.0 Xml;HDR=YES'";

                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    try
                    {
                        connection.Open();
                        //this next line assumes that the file is in default Excel format with Sheet1 as the first sheet name, adjust accordingly
                        OleDbCommand cmd = new OleDbCommand("INSERT INTO [Sheet1$](ID, Anrede, Nachname, Vorname, Addresse, PLZ, Ort, Telefon, Email, Betrag) VALUES(@ID, @Anrede, @Nachname, @Vorname, @Addresse, @PLZ, @Ort, @Telefon, @Email, @Betrag)", connection);

                        if (connection.State == ConnectionState.Open)
                            {
                                cmd.Parameters.Add("@ID", OleDbType.VarChar, 50).Value =ID;
                                cmd.Parameters.Add("@Anrede", OleDbType.VarChar, 50).Value =ID;
                                cmd.Parameters.Add("@Nachname", OleDbType.VarChar, 50).Value =ID;
                                cmd.Parameters.Add("@Vorname", OleDbType.VarChar, 50).Value =ID;
                                cmd.Parameters.Add("@Addresse", OleDbType.VarChar, 50).Value =ID;
                                cmd.Parameters.Add("@PLZ", OleDbType.VarChar, 50).Value =ID;
                                cmd.Parameters.Add("@Ort", OleDbType.VarChar, 50).Value =ID;
                                cmd.Parameters.Add("@Telefon", OleDbType.VarChar, 50).Value =ID;
                                cmd.Parameters.Add("@Email", OleDbType.VarChar, 50).Value =ID;
                                cmd.Parameters.Add("@Betrag", OleDbType.VarChar, 50).Value =ID;
                            try
                                {
                                    cmd.ExecuteNonQuery();
                                    MessageBox.Show("DATA ADDED");
                                    connection.Close();
                                }
                            catch (OleDbException expe)
                                {
                                    MessageBox.Show(expe.Message);
                                    connection.Close();
                                }
                            }
                    }

                    catch (OleDbException)
                    {
                        Console.WriteLine("Something went wrong with DB connecting for addin new member");
                    }
                }
            }
        }

        public void deleteMember(Person person)
        {
            String ID = person.ID;
            dbPath = MemberAdministration.Properties.Settings.Default.dbPath;
            if (dbPath != null)
            {
                /*
                 * Connection begins here when dbPath is set   ********   *********   *********
                 */
                String connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + dbPath + "; Extended Properties='Excel 12.0 Xml;HDR=YES'";

                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    try
                    {
                        connection.Open();
                        //this next line assumes that the file is in default Excel format with Sheet1 as the first sheet name, adjust accordingly
                        OleDbCommand cmd = new OleDbCommand("DELETE FROM [Sheet1$] * WHERE ID = @ID", connection);

                        if (connection.State == ConnectionState.Open)
                        {
                            cmd.Parameters.Add("@ID", OleDbType.VarChar, 50).Value = ID;
                            try
                            {
                                cmd.ExecuteNonQuery();
                                MessageBox.Show("Person :"+person.Name+" Deleted");
                                connection.Close();
                            }
                            catch (OleDbException expe)
                            {
                                MessageBox.Show(expe.Message);
                                connection.Close();
                            }
                        }
                    }

                    catch (OleDbException)
                    {
                        Console.WriteLine("Something went wrong with DB connecting for addin new member");
                    }
                }
            }
        }
    }
}
