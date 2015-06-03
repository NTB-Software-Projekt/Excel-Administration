using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MemberAdministration
{
    /// <summary>
    /// All CRUD operations are handled in this class.
    /// Connection to the Excel File is managed trough the OleDB Connector with an ISAM driver.
    /// We need to set the Provider String with the path to the File for the connection to be established.
    /// </summary>
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
        /// <returns> DataSet with all the Member Data </returns>
        public DataSet getAllMembers()
        {
            dbPath = MemberAdministration.Properties.Settings.Default.dbPath;
            if (dbPath != null)
            {
                /*
                 * Connection begins here when dbPath is set   ********   *********   *********
                 * 
                 * 
                 * This is the Provider String needed for the driver to work.
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

        /// <summary>
        /// Pulling a specific member from the Database, filtered by the inputted name,surname, address or state. 
        /// </summary>
        /// <param name="member">name, surname, address or state of the wanted person.</param>
        /// <returns> DataSet with the specific person information.</returns>
        public DataSet getThisMember(String member)
        {
            dbPath = MemberAdministration.Properties.Settings.Default.dbPath;
            if (dbPath != null)
            {
                /*
                 * Connection begins here when dbPath is set   ********   *********   *********
                 * 
                 * 
                 * This is the Provider String needed for the driver to work.
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

        /// <summary>
        /// Updating the existing information of the Member in the DB with the new ones taken from the Form
        /// </summary>
        /// <param name="person">Person object containing all the needed informatino</param>
        public void updateMember(Person person)
        {
            String ID = person.ID;
            String title = person.Title;
            String surname = person.Surname;
            String name = person.Name;
            String address = person.Address;
            String plz = person.Plz;
            String state = person.State;
            String telephone = person.Telephone.ToString();
            String mail = person.Mail;
            String amount = person.Amount.ToString();

            dbPath = MemberAdministration.Properties.Settings.Default.dbPath;
            if (dbPath != null)
            {
                /*
                 * Connection begins here when dbPath is set   ********   *********   *********
                 * 
                 * 
                 * This is the Provider String needed for the driver to work.
                 */
                String connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + dbPath + "; Extended Properties='Excel 12.0 Xml;HDR=YES'";

                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    try
                    {
                        connection.Open();
                        //this next line assumes that the file is in default Excel format with Sheet1 as the first sheet name, adjust accordingly
                        OleDbCommand cmd = new OleDbCommand("UPDATE [Sheet1$] SET Anrede = ?, Nachname = ?, Vorname = ?, Adresse = ?, PLZ = ?, Ort = ?, Telefon = ?, EMail = ?, Betrag = ? WHERE ID = ?", connection);
                        if (connection.State == ConnectionState.Open)
                        {
                            cmd.Parameters.Add("@Anrede", OleDbType.VarChar, 50).Value = title;
                            cmd.Parameters.Add("@Nachname", OleDbType.VarChar, 50).Value = surname;
                            cmd.Parameters.Add("@Vorname", OleDbType.VarChar, 50).Value = name;
                            cmd.Parameters.Add("@Adresse", OleDbType.VarChar, 50).Value = address;
                            cmd.Parameters.Add("@PLZ", OleDbType.VarChar, 50).Value = plz;
                            cmd.Parameters.Add("@Ort", OleDbType.VarChar, 50).Value = state;
                            cmd.Parameters.Add("@Telefon", OleDbType.VarChar, 50).Value = telephone;
                            cmd.Parameters.Add("@Email", OleDbType.VarChar, 50).Value = mail;
                            cmd.Parameters.Add("@Betrag", OleDbType.VarChar, 50).Value = amount;
                            cmd.Parameters.Add("@ID", OleDbType.VarChar, 50).Value = ID;
                            try
                            {
                                cmd.ExecuteNonQuery();
                                MessageBox.Show("Person data of " + person.Name + " UPDATED");
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
                        Console.WriteLine("Something went wrong with DB connecting for Updating");
                    }
                }
            }
        }

        /// <summary>
        /// Inserting the new person into the DB. We strip all the information from the Person object and pass it to the query
        /// </summary>
        /// <param name="person">Person object containing all the needed informatino</param>
        public void insertNewMember(Person person)
        {
            String ID = person.ID;
            String title = person.Title;
            String surname = person.Surname;
            String name = person.Name;
            String address = person.Address;
            String plz = person.Plz;
            String state = person.State;
            String telephone = person.Telephone.ToString();
            String mail = person.Mail;
            String amount = person.Amount.ToString();

            dbPath = MemberAdministration.Properties.Settings.Default.dbPath;
            if (dbPath != null)
            {
                /*
                 * Connection begins here when dbPath is set   ********   *********   *********
                 * 
                 * 
                 * This is the Provider String needed for the driver to work.
                 */
                String connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + dbPath + "; Extended Properties='Excel 12.0 Xml;HDR=YES'";
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    try
                    {
                        connection.Open();
                        //this next line assumes that the file is in default Excel format with Sheet1 as the first sheet name, adjust accordingly
                        OleDbCommand cmd = new OleDbCommand("INSERT INTO [Sheet1$](ID, Anrede, Nachname, Vorname, Adresse, PLZ, Ort, Telefon, EMail, Betrag) VALUES(@ID, @Anrede, @Nachname, @Vorname, @Adresse, @PLZ, @Ort, @Telefon, @EMail, @Betrag)", connection);
                        if (connection.State == ConnectionState.Open)
                        {
                            cmd.Parameters.Add("@ID", OleDbType.VarChar, 50).Value = ID;
                            cmd.Parameters.Add("@Anrede", OleDbType.VarChar, 50).Value = title;
                            cmd.Parameters.Add("@Nachname", OleDbType.VarChar, 50).Value = surname;
                            cmd.Parameters.Add("@Vorname", OleDbType.VarChar, 50).Value = name;
                            cmd.Parameters.Add("@Adresse", OleDbType.VarChar, 50).Value = address;
                            cmd.Parameters.Add("@PLZ", OleDbType.VarChar, 50).Value = plz;
                            cmd.Parameters.Add("@Ort", OleDbType.VarChar, 50).Value = state;
                            cmd.Parameters.Add("@Telefon", OleDbType.VarChar, 50).Value = telephone;
                            cmd.Parameters.Add("@EMail", OleDbType.VarChar, 50).Value = mail;
                            cmd.Parameters.Add("@Betrag", OleDbType.VarChar, 50).Value = amount;
                            try
                            {
                                cmd.ExecuteNonQuery();
                                MessageBox.Show("Person " + person.Name + " ADDED");
                                connection.Close();
                            }
                            catch (OleDbException expe)
                            {
                                MessageBox.Show(expe.Message);
                                connection.Close();
                            }
                        }
                        connection.Open();
                        cmd = new OleDbCommand("INSERT INTO [Sheet1$](ID, Anrede, Nachname, Vorname, Adresse, PLZ, Ort, Telefon, EMail, Betrag) VALUES(@ID, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL)", connection);
                        if (connection.State == ConnectionState.Open)
                        {
                            Int32 nextID = Int32.Parse(ID) + 1;
                            ID = nextID.ToString();
                            cmd.Parameters.Add("@ID", OleDbType.VarChar, 50).Value = ID;

                            try
                            {
                                cmd.ExecuteNonQuery();
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

        /// <summary>
        /// Deleting one specific member, filtered by the personID
        /// </summary>
        /// <param name="personID">String with the wanted ID</param>
        public void deleteMember(String personID)
        {
            String checkID = personID;
            dbPath = MemberAdministration.Properties.Settings.Default.dbPath;
            if (dbPath != null)
            {
                /*
                 * Connection begins here when dbPath is set   ********   *********   *********
                 * 
                 * 
                 * This is the Provider String needed for the driver to work.
                 */
                String connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + dbPath + "; Extended Properties='Excel 12.0 Xml;HDR=YES'";
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    try
                    {
                        connection.Open();
                        //this next line assumes that the file is in default Excel format with Sheet1 as the first sheet name, adjust accordingly
                        OleDbCommand cmd = new OleDbCommand("UPDATE [Sheet1$] SET Anrede = NULL, Nachname = NULL, Vorname = NULL, Adresse = NULL, PLZ = NULL, Ort = NULL, Telefon = NULL, EMail = NULL, Betrag = NULL WHERE ID = ?", connection);
                        if (connection.State == ConnectionState.Open)
                        {
                            cmd.Parameters.Add("@ID", OleDbType.VarChar, 50).Value = checkID;
                            try
                            {
                                cmd.ExecuteNonQuery();
                                MessageBox.Show("Person DELETED");
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
                        Console.WriteLine("Something went wrong with DB connecting for Updating");
                    }
                }
            }
        }

        /// <summary>
        /// Getting the max (latest) ID from the Excel Database column "ID"
        /// </summary>
        /// <returns>String containing the maximum ID</returns>
        public String maxID()
        {
            String maxID;
            dbPath = MemberAdministration.Properties.Settings.Default.dbPath;
            if (dbPath != null)
            {
                /*
                 * Connection begins here when dbPath is set   ********   *********   *********
                 * 
                 * 
                 * This is the Provider String needed for the driver to work.
                 */
                String connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + dbPath + "; Extended Properties='Excel 12.0 Xml;HDR=YES'";
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    try
                    {
                        //this next line assumes that the file is in default Excel format with Sheet1 as the first sheet name, adjust accordingly
                        OleDbCommand cmd = new OleDbCommand("SELECT MAX(ID) FROM [Sheet1$]", connection);
                        connection.Open();

                        if (connection.State == ConnectionState.Open)
                        {
                            try
                            {
                                using (OleDbDataReader reader = cmd.ExecuteReader())
                                {

                                    while (reader.Read())
                                    {
                                        maxID = reader.GetValue(0).ToString();
                                        return maxID;
                                    }
                                }
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
                        Console.WriteLine("Something went wrong with DB connection");
                    }
                }
            }
            // This should never be returned, but if it does, the least amount of damage is done.
            return "500";
        }
    }
}
