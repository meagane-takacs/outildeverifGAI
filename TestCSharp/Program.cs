using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace Nmx.DataEngineering.EngineeringConfigurationScript
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.Linq;
    using Nmx.Public;
    using Nmx.Public.DE;
    using System.IO;
    using System.Text;
    using System.Reflection;




    public class EngineeringConfigurationScript : EngineeringConfigurationScriptBase
    {
        //** Donnees membres ** 

        //dossier courant
        static string DIRECTORY_CURRENT = "";
        static string DIRECTORY_LOG = "";
        //Certains fichiers a lire ou a exclure (utilisateurs qui ne doivent pas y avoir acces)
        static string DIRECTORY_PARAMETRE = "";
        //erreur,type d'erreur
        static int RETC = 0;
        static int NBR_JOUR_PURG_LOG = 5;
        static DateTime DTNOW;
        static string DATENOW;
        static string TIMENOW;
        static string DATELOG;
        static string MESSAGE_LOG = "";
        static string TYPE_MESSAGE = "Inform";
        static string PROGRAM_NAME = "Gai_verify";
        static string AUTHNAME = "Takacs_Meagane";
        static string COMPUTERNAME = "";
        static string USERPROFILE = "";
        static string FILELOG = "";
        static bool WRITECONSOLEOUTPUT = true;




        public override void Execute(WorkspaceObject[] wsObjects)
        {
            //Recuperation de l'endroit ou se trouve le dossier courant
            DIRECTORY_CURRENT = System.IO.Directory.GetCurrentDirectory();
            //On renseigne Directory_log et directory_parametre sous forme de chaine de caractere qui renseigne les chemin des dossiers
            DIRECTORY_LOG = string.Format("C:\\Users\\Public\\Documents\\RTE\\Log_Meagane");
            DIRECTORY_PARAMETRE = string.Format("{0}\\Parametre", DIRECTORY_CURRENT);




            /**************************************************************************************************************************************/
            /* Gestion des log ancienne Recherche des fichiers log  pour menage                                                                   */
            /**************************************************************************************************************************************/

            //Creation d'une nouvelle instance classe DirectoryInfo => expose des m?thodes d'instance pour creer, se deplacer dans et enumerer des repertoires et sous repertoires
            //DIR_LOG est un objet de type directory info => on apelle le constructeur de directoryinfo et en parametre on lui passe un path (car le contructeur marche comme ?a d'apres la doc)
            DirectoryInfo DIR_LOG = new DirectoryInfo(DIRECTORY_LOG);
            //Si le repertoire directory log n'existe pas, on le creer
            if (!Directory.Exists(DIRECTORY_LOG))
            {
                DirectoryInfo DIR_LOG_CREAT = Directory.CreateDirectory(DIRECTORY_LOG);
            }
            //Recuperation de tous les fichiers d'extensions .log (Classe FileInfo => Fournit des proprietes et des methodes d'instance pour creer, copier, supprimer deplacer et ouvrir des fichiers, facilite la creations d'objets file stream)
            // FILELOG_OLD est un objet de type FileInfo dans un tableau. On retourne la liste des fichiers du repertoire actuel correspondant au modele de recherche donne (donc .log)
            FileInfo[] FILELOG_OLD = DIR_LOG.GetFiles("*.log");
            //Pour chaque fichiers .log
            foreach (FileInfo FILELOG_OLD_LIST in FILELOG_OLD)
            {
                //Recuperation du chemin entier et de la date du fichier courant
                string FILENAME_LOG_OLD = FILELOG_OLD_LIST.FullName;
                DateTime DATE_CREAT_FILES = FILELOG_OLD_LIST.CreationTime.Date;
                //Affichage du nom du fichier courant
                System.Console.WriteLine(" filelog ancienne " + FILENAME_LOG_OLD);
                //Si la date de creation du courant date de 5 jours, on l'efface
                if (DATE_CREAT_FILES < DateTime.Now.AddDays(-NBR_JOUR_PURG_LOG))
                {
                    File.Delete(string.Format("{0}", FILENAME_LOG_OLD));
                }
            }
            //Recuperation dans l'environnement de la variable computername
            COMPUTERNAME = Environment.GetEnvironmentVariable("computername");




            /********************************************************************************************************/
            /*  Affectation des dates pour les fichiers de log                                                      */
            /********************************************************************************************************/

            //Construction d'une date
            DateTime DTNOW = DateTime.Now;
            //On lui affecte ce format la (premier chiffre correspondent a l'index, les second apres les : sont le nombres de chiffre qu'on veut)
            DATENOW = string.Format("{0}{1:00}{2:00}", DTNOW.Year, DTNOW.Month, DTNOW.Day);
            //On definit l'h les min et les sec dans des objets differents pour pouvoir les mettre dans le nom du filelog sans embeter avec les ":"
            string HNOW = string.Format("{0:00}", DTNOW.Hour);
            string MNOW = string.Format("{0:00}", DTNOW.Minute);
            string SNOW = string.Format("{0:00}", DTNOW.Second);
            // dd: jours sur deux chiffres
            DATELOG = string.Format("{0:dd}", DTNOW);
            FILELOG = string.Format("{0}\\{1}_{2}_{3}{4}{5}.log", DIRECTORY_LOG, PROGRAM_NAME, DATENOW, HNOW, MNOW, SNOW);
            MESSAGE_LOG = string.Format("Demarrage application {0} Auteur={1} Computer_name={2}", PROGRAM_NAME, AUTHNAME, COMPUTERNAME);
            Writelog(FILELOG, MESSAGE_LOG, TYPE_MESSAGE);




            /*************************************************************************************************************************/
            /* insertion du code                                                                                                     */
            /*************************************************************************************************************************/

            AnalyserFichiersRef();


        } //Fin de public override Execute




        /*********************************************************************************************************************************/
        /*  methods                                                                                                                      */
        /*********************************************************************************************************************************/

        private void AnalyserFichiersRef()
        {
            // Recuperer tous les fichiers .ref
            // On declare DIRECTORY_REF vide
            string DIRECTORY_REF = "";
            //On recupere le nom de l'environnement du user profile
            USERPROFILE = Environment.GetEnvironmentVariable("USERPROFILE");
            //String.Format method convertit la valeur des objets en chainesselon les formats specifies et les insere dans une autre chaine. On insere USERPROFILE dans le {0}
            //On affecte a DIRECTORY_REF une valeur formatee
            DIRECTORY_REF = string.Format("{0}\\AppData\\Roaming\\ABB\\PED500\\pic", USERPROFILE);
            if (!Directory.Exists(DIRECTORY_REF))
            {
                Errors.ReportStatus("Dossier pic inexistant");
                MESSAGE_LOG = string.Format("Dossier pic inexistant");
                Writelog(FILELOG, MESSAGE_LOG, TYPE_MESSAGE);

            }

            //Creation d'une nouvelle instance classe DirectoryInfo => expose des m?thodes d'instance pour creer, se deplacer dans et enumerer des repertoires et sous repertoires
            //DIR_LOG est un objet de type directory info => on apelle le constructeur de directoryinfo et en parametre on lui passe un path (car le contructeur marche comme ?a d'apres la doc)
            DirectoryInfo DIR_REF = new DirectoryInfo(DIRECTORY_REF);

            //Recuperation de tous les fichiers d'extensions .ref (Classe FileInfo => Fournit des proprietes et des methodes d'instance pour creer, copier, supprimer deplacer et ouvrir des fichiers, facilite la creations d'objets file stream)
            // FILELOG_OLD est un objet de type FileInfo dans un tableau. On retourne la liste des fichiers du repertoire actuel correspondant au modele de recherche donne (donc .ref)
            FileInfo[] FILEREF_LIST = DIR_REF.GetFiles("*.ref");
            //Pour chaque fichiers .ref
            foreach (FileInfo FILEREF_CURRENT in FILEREF_LIST)
            {
                MESSAGE_LOG = string.Format("Appel de la methode AnalyserRefFile {0}", FILEREF_CURRENT.FullName);
                Writelog(FILELOG, MESSAGE_LOG, TYPE_MESSAGE);

                //Liste des elements present dans le .ref que l'on range dans myElementsRefPart sur laquelle on appelle AnalyserRefFile qui nous permettra d'en extraire l'ID et le Pres_ext_ID
                List<ElemRefPart> myElementsRefPart;
                myElementsRefPart = AnalyserRefFile(FILEREF_CURRENT);

                //preparation de le log
                MESSAGE_LOG = string.Format("Fin de AnalyserRefFile {0}", FILEREF_CURRENT.FullName);
                Writelog(FILELOG, MESSAGE_LOG, TYPE_MESSAGE);

                //Appel de la methode GetMeasurandBdd avec en parametre la liste des elements present dans le .ref
                GetMeasurandBdd(myElementsRefPart);
            }

        }




        // ** Methode qui analyse les fichiers ref **
        //List => Repr?sente une liste typees d'objet accessible par index. On fait une liste d'elemRefPart. On passe en parametre un Reffile de type FileInfo
        private List<ElemRefPart> AnalyserRefFile(FileInfo RefFile)
        {

            List<ElemRefPart> myElementsRefPart = new List<ElemRefPart>();
            ElemRefPart myElement = new ElemRefPart();
            int counter = 0;
            string line;
            char[] charSeparators = new char[] { ':' };

            //StreamReader implemente textreader => lecteur de caractere capable de lire une s?rie de caractere (ici on s'en sert pour lire dans le fichier)
            // (File => fournit des methodes statiques pour creer copier supprimer deplacer et ouvrir un fichier unique)
            StreamReader File = new StreamReader(RefFile.FullName);
            //Tant qu'il y a des lignes a lire (que la ligne suivante du fichier n'est pas nulle)
            while ((line = File.ReadLine()) != null)
            {
                //Split => retourne un tableau de chaine qui contient les sous chaine de cette instance, separees par les elements d'une chaine 
                // ici on divise une chaine en un nombre max de sous chaine d'apres le separateur de caractere fourni (2 => nb maximal d'el?ments attendus dans le tableau)
                string[] tokens = line.Split(charSeparators, 2, StringSplitOptions.None);
                //Enlever les espaces en debut et fin (tokens[0] = index 0 donc la key et tokens[1] la valeur qu'on a s?parer avec le split)
                tokens[0] = tokens[0].Trim();
                tokens[1] = tokens[1].Trim();

                //En fonction de la valeur de token[0]
                switch (tokens[0])
                {
                    //Si ca vaut ca
                    case "begin_elem_ref_part":
                        //Je creer un nouvel element courant => a chaque begin elem ref part, je creer un elemrefpart
                        myElement = new ElemRefPart();
                        break;

                    //Si ca vaut ca
                    case "element_id":
                        //Je fixe la valeur de element id dans my element = Integer derriere car conversion chaine de caractere vers entier car setElement_id attend un entier et et tokens[1] est une chaine de caractere
                        myElement.Element_id = int.Parse(tokens[1]);
                        break;

                    //si ca vaut ca
                    case "pres_ext_identity":
                        //Je fixe la valeur de pres_ext_identity dans my element avec pour parametre (tokens[1])
                        myElement.Pres_ext_identity = tokens[1];
                        //On enleve les guillemets
                        myElement.Pres_ext_identity = myElement.Pres_ext_identity.Substring(1, myElement.Pres_ext_identity.Length - 2);
                        break;

                    case "end_elem_ref_part":
                        myElementsRefPart.Add(myElement);
                        break;

                }

                //System.out.println(tokens[0]+tokens[1]);
            }
            File.Close();

            EcrireFichierResultat(RefFile.Name, myElementsRefPart);

            //Retour pour pouvoir le passer a une autre methode ensuite
            return myElementsRefPart;
        }





        // ** Methode qui permet d'ecrire les resultat d'AnalyserRefFile dans un fichier **
        private void EcrireFichierResultat(string pNomRef, List<ElemRefPart> pMyElementsRefPart)
        {
            //Ecriture dans un fichier csv
            string text = "ID;PRES_EXTERNAL_ID\n";
            string USERNAME = Environment.GetEnvironmentVariable("USERNAME");
            string docPath = string.Format("C:\\Users\\Public\\Documents\\RTE\\CSV_{0}", USERNAME);
            if (!Directory.Exists(docPath))
            {
                DirectoryInfo DIR_DOCPATH_CREAT = Directory.CreateDirectory(docPath);
            }
            string fileName = string.Format("{0}.csv", pNomRef);
            File.WriteAllText(Path.Combine(docPath, fileName), text);

            //Ecriture dans le log
            MESSAGE_LOG = string.Format("Appel de la methode EcrireFichierResultat {0}\\{1}", docPath, fileName);
            Writelog(FILELOG, MESSAGE_LOG, TYPE_MESSAGE);

            for (int i = 0; i < pMyElementsRefPart.Count; i++)
            {
                //on affiche l' ID et le pres_external_identity dans le log
                Console.WriteLine("ID :" + pMyElementsRefPart[i].Element_id + " " + "PRES_EXTERNAL_ID : " + pMyElementsRefPart[i].Pres_ext_identity);
                text = pMyElementsRefPart[i].Element_id + ";" + pMyElementsRefPart[i].Pres_ext_identity + "\n";
                File.AppendAllText(Path.Combine(docPath, fileName), text);

            }
            //Ecriture dans le log
            MESSAGE_LOG = string.Format("Fin de EcrireFichierResultat {0}\\{1}", docPath, fileName);
            Writelog(FILELOG, MESSAGE_LOG, TYPE_MESSAGE);
        }




        // ** Methode qui permet de rajouter un espace avant chaque majusucule les types dans object list (get definition) et le type dans la BDD ne correspondent pas **
        private string AddSpacesToSentence(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return "";
            StringBuilder newText = new StringBuilder(text.Length * 2);
            newText.Append(text[0]);
            for (int i = 1; i < text.Length; i++)
            {
                if (char.IsUpper(text[i]) && text[i - 1] != ' ')
                    newText.Append(' ');
                newText.Append(text[i]);
            }
            return newText.ToString();
        }




        // ** Methode qui permet de recuperer les donnees dans la BDD avec pour parametre les resultats d'analyser Ref part (tout les ID et pres ext ID) **

        private void GetMeasurandBdd(List<ElemRefPart> pMyElementsRefPart)
        {
            //Ouverture du dossier et preparation, ecrasement avec un texte vide
            string text = "EXTERNAL_ID;TYPE;STATION;BAY;SUBNET;EQUIPEMENT;EQUIPEMENT_TYPE;ZONE\n";
            string USERNAME = Environment.GetEnvironmentVariable("USERNAME");
            string docPath = string.Format("C:\\Users\\Public\\Documents\\RTE\\GetMeasurand_{0}", USERNAME);
            if (!Directory.Exists(docPath))
            {
                DirectoryInfo DIR_DOCPATH_CREAT = Directory.CreateDirectory(docPath);
            }
            string fileName = string.Format("GetMesurandBdd.csv");
            File.WriteAllText(Path.Combine(docPath, fileName), text);

            MESSAGE_LOG = string.Format("Debut de GetMeasurandBdd");
            Writelog(FILELOG, MESSAGE_LOG, TYPE_MESSAGE);

            //On instancie de nouveau workspace objects
            IList<WorkspaceObject> measurands = new List<WorkspaceObject>();
            IList<WorkspaceObject> indications = new List<WorkspaceObject>();
            IList<WorkspaceObject> processValue = new List<WorkspaceObject>();
            //Pour stocker le resultat des select sur la base
            IList<WorkspaceObject> objectsList = new List<WorkspaceObject>();

            //for (int i = 0; i < pMyElementsRefPart.Count; i++)
            //Pour [i] de 0 ? 10
            for (int i = 510; i < 520; i++)
            //for (int i = 0; i < 10; i++)
            {
                //Pour chaque element du tableau pMyElementRefPart, je prend son external ID
                string lExtIdentity = pMyElementsRefPart[i].Pres_ext_identity;
                if (WRITECONSOLEOUTPUT == true)
                {
                    Errors.ReportStatus(" ");
                    Errors.ReportStatus("ExtIdentity : " + lExtIdentity);
                }
                //ecriture dans le log
                MESSAGE_LOG = string.Format("ExtIdentity : " + lExtIdentity);
                Writelog(FILELOG, MESSAGE_LOG, TYPE_MESSAGE);

                //Recuperation de tout les membres du types objects types
                MemberInfo[] members = typeof(ObjectTypes).GetMembers();

                //string[] mesTypes = { "measurand", "Process Value", "Indication" };

                foreach (MemberInfo m in members)
                {
                    string lTypeCourant = AddSpacesToSentence(m.Name);

                    //Recuperation des objets du type courant ou leur external ID = au external ID courant. On parcourt tout les types possible pour le type courant on regarde s'il y  a un objet avec le ext ID courant
                    objectsList = Workspace.GetObjects(lTypeCourant).Where(o => GetExtIdentity(o) == lExtIdentity).ToList();
                    if (objectsList.Count == 1)
                    {
                        if (WRITECONSOLEOUTPUT == true)
                        {
                            Errors.ReportStatus(" ");
                            Errors.ReportStatus("Objet de type : " + lTypeCourant);
                        }
                        //ecriture dans le log
                        MESSAGE_LOG = string.Format("Objet de type : " + lTypeCourant);
                        Writelog(FILELOG, MESSAGE_LOG, TYPE_MESSAGE);

                        // Affichage station
                        if (objectsList[0].Station != null)
                        {
                            if (WRITECONSOLEOUTPUT == true)
                            {
                                Errors.ReportStatus("Station : " + objectsList[0].Station.Name);
                            }
                            //ecriture dans le log
                            MESSAGE_LOG = string.Format("Station : " + objectsList[0].Station.Name);
                            Writelog(FILELOG, MESSAGE_LOG, TYPE_MESSAGE);
                        }
                        else
                        {
                            if (WRITECONSOLEOUTPUT == true)
                            {
                                Errors.ReportStatus("Pas de station associee");
                            }
                        }

                        //affichage zone
                        string zone = "Pas de zone";
                        string zoneName = "";
                        //On recupere un objet Station identifie par son nom
                        if (objectsList[0].Station != null)
                        {
                            string lStationName = objectsList[0].Station.Name;

                            if (lStationName != "")
                            {
                                WorkspaceObject lStationObj = GetStationObject(lStationName);

                                zone = lStationObj.GetValue<string>("ZONE_REF");
                                zoneName = GetZone(zone);
                                if (WRITECONSOLEOUTPUT == true)
                                {
                                    Errors.ReportStatus("zone : " + zoneName);
                                }
                            }
                        }

                        //Affichage bay
                        string bay = objectsList[0].GetValue<string>("BAY_REF");
                        //On apelle la methode get bay name pour recuperer le nom de l'objet et faire le lien entre les tables
                        string bayName = GetBayName(bay);
                        if (WRITECONSOLEOUTPUT == true)
                        {
                            Errors.ReportStatus("Bay : " + bayName);
                        }
                        //ecriture log
                        MESSAGE_LOG = string.Format("Bay : " + bayName);
                        Writelog(FILELOG, MESSAGE_LOG, TYPE_MESSAGE);

                        //affichage subnet 
                        string subnet = objectsList[0].GetValue<string>("SUBNET_REF");
                        //On apelle la methode get subnet name pour recuperer le nom de l'objet et faire le lien entre les tables
                        string subnetName = GetSubnetName(subnet);
                        if (WRITECONSOLEOUTPUT == true)
                        {
                            Errors.ReportStatus("Subnet : " + subnetName);
                        }
                        //ecriture log
                        MESSAGE_LOG = string.Format("Subnet : " + subnetName);
                        Writelog(FILELOG, MESSAGE_LOG, TYPE_MESSAGE);

                        //affichage equipment
                        string equipment = objectsList[0].GetValue<string>("EQUIPMENT_REF");
                        //On apelle la methode get equpment name pour recuperer le nom de l'objet et faire le lien entre les tables
                        string equipmentName = GetEquipmentName(equipment);
                        if (WRITECONSOLEOUTPUT == true)
                        {
                            Errors.ReportStatus("Equipement : " + equipmentName);
                        }
                        //ecriture log
                        MESSAGE_LOG = string.Format("Equipement : " + equipmentName);
                        Writelog(FILELOG, MESSAGE_LOG, TYPE_MESSAGE);

                        // type equipement
                        string typeEquipment = GetEquipmentType(equipment);
                        if (WRITECONSOLEOUTPUT == true)
                        {
                            Errors.ReportStatus("TypeEquipment : " + typeEquipment);
                        }
                        //ecriture log
                        MESSAGE_LOG = string.Format("TypeEquipment : " + typeEquipment);
                        Writelog(FILELOG, MESSAGE_LOG, TYPE_MESSAGE);

                        //Ecriture dans csv
                        if (objectsList[0].Station != null)
                        {
                            text = pMyElementsRefPart[i].Pres_ext_identity + ";" + lTypeCourant + ";" + objectsList[0].Station.Name + ";" + bayName + ";" + subnetName + ";" + equipmentName + ";" + typeEquipment + ";" + zoneName + "\n";
                        }
                        else
                        {
                            text = pMyElementsRefPart[i].Pres_ext_identity + ";" + lTypeCourant + ";" + "Pas de station" + ";" + bayName + ";" + subnetName + ";" + equipmentName + ";" + typeEquipment + ";" + zoneName + "\n";
                        }
                        File.AppendAllText(Path.Combine(docPath, fileName), text);


                        //break; // On a trouvz notre type, pas besoin de regarder les autres types --> on sort du foreach ( MemberInfo m in members)
                        // NON : sinon on ne voit pas les extIdentity qui renvoi plusieurs objets de type differents
                    }
                    if (objectsList.Count > 1)
                    {
                        if (WRITECONSOLEOUTPUT == true)
                        {
                            Errors.ReportStatus("J'ai trouve PLUSIEURS objet de type : " + lTypeCourant);
                        }
                    }

                }
            }
            MESSAGE_LOG = string.Format("Fin de GetMeasurandBdd");
            Writelog(FILELOG, MESSAGE_LOG, TYPE_MESSAGE);
        }




        // ** Methode qui recupere le nom de l'equipement via son ID lorsque l'ID parametre rencontre l'ID de la BDD **
        private string GetEquipmentName(string ID)
        {
            //Recuperation de tout les membres du types objectstypes (on ne sait pas de quel type est equipment)
            MemberInfo[] members = typeof(ObjectTypes).GetMembers();

            IList<WorkspaceObject> objectsList = new List<WorkspaceObject>();
            foreach (MemberInfo m in members)
            {
                //espace a chaque majusucule (sauf la premiere)
                string lTypeCourant = AddSpacesToSentence(m.Name);
                //Recuperation des objets du type courant ou leur external ID = au external ID courant
                objectsList = Workspace.GetObjects(lTypeCourant).Where(o => GetObjectID(o) == ID).ToList();
                if (objectsList.Count == 1)
                {
                    return objectsList[0].Name;
                }
            }
            return "Pas d'equipement";
        }

        private string GetEquipmentType(string ID)
        {
            MemberInfo[] members = typeof(ObjectTypes).GetMembers();
            IList<WorkspaceObject> objectsList = new List<WorkspaceObject>();
            foreach (MemberInfo m in members)
            {
                //espace a chaque majusucule (sauf la premiere)
                string lTypeCourant = AddSpacesToSentence(m.Name);
                //Recuperation des objets du type courant ou leur external ID = au external ID courant
                objectsList = Workspace.GetObjects(lTypeCourant).Where(o => GetObjectID(o) == ID).ToList();
                if (objectsList.Count == 1)
                {
                    return objectsList[0].TypeName;
                }
            }

            return "Pas de type equipment";
        }

        // ** Methode qui recupere le nom du subnet via son ID lorsque l'ID parametre rencontre l'ID de la BDD **

        private string GetSubnetName(string ID)
        {
            IList<WorkspaceObject> objectsList = new List<WorkspaceObject>();
            objectsList = Workspace.GetObjects("subnet").Where(o => GetObjectID(o) == ID).ToList();
            if (objectsList.Count == 1)
            {
                return objectsList[0].Name;
            }
            return "Pas de subnet";
        }




        // ** Methode qui recupere le nom de la bay via son ID lorsque l'ID parametre rencontre l'ID de la BDD **

        private string GetBayName(string ID)
        {
            IList<WorkspaceObject> objectsList = new List<WorkspaceObject>();
            objectsList = Workspace.GetObjects("bay").Where(o => GetObjectID(o) == ID).ToList();
            if (objectsList.Count == 1)
            {
                return objectsList[0].Name;
            }
            return "Pas de bay";
        }



        // ** Methode qui recupere le nom de l'objet ID et le transforme en string **

        private string GetObjectID(WorkspaceObject obj)
        {
            return obj.Id.ToString();
        }



        // ** Methode qui recupere la valeur de la propriete de l'objet external identity **

        private string GetExtIdentity(WorkspaceObject obj)
        {
            return obj.GetValue<string>("EXTERNAL_IDENTITY");
        }

        private WorkspaceObject GetStationObject(string StationName)
        {
            IList<WorkspaceObject> objectsList = new List<WorkspaceObject>();
            objectsList = Workspace.GetObjects("Station").Where(o => o.Name == StationName).ToList();

            if (objectsList.Count == 1)
            {
                return objectsList[0];
            }

            return null;
        }

        private string GetZone(string ID)
        {
            //Errors.ReportStatus("GetZone de" + ID);

            IList<WorkspaceObject> objectsList = new List<WorkspaceObject>();
            objectsList = Workspace.GetObjects("zone").Where(o => GetObjectID(o) == ID).ToList();

            //Errors.ReportStatus("GetZone de" + ID + " recuperation de nbobjet : "+ objectsList.Count);

            if (objectsList.Count == 1)
            {
                return objectsList[0].Name;
            }
            return "Pas de zone";
        }


        /*********************************************************************************************************************************/
        /*  ecriture de le log                                                                                                           */
        /*********************************************************************************************************************************/

        private static void Writelog(string FILELOG, string MESSAGE_LOG, string TYPE_MESSAGE)
        {
            //On construit une date 
            DateTime DTNow = DateTime.Now;
            //On declare
            String Message_Log;
            //On declare
            DateTime dtCurr;
            dtCurr = DateTime.Now;
            //On affecte
            Message_Log = String.Format("{0}-{1:00}-{2:00} {3:00}:{4:00}:{5:00} {6,8} {7}", dtCurr.Year, dtCurr.Month, dtCurr.Day, dtCurr.Hour, dtCurr.Minute, dtCurr.Second, TYPE_MESSAGE, MESSAGE_LOG);
            // Stream writer implemente TextWriter pour ecrire les caracteres dans un flux selon un encodage particulier. File.Append creer un element streamwriter qui ajoute du texte encode en UTF8
            using (StreamWriter sw = File.AppendText(FILELOG))
            {
                try
                {
                    //On appelle la methode WriteLine et on lui passe en parametre message log. Elle ecrira la chaine de caractere de message log.
                    sw.WriteLine(Message_Log);
                }
                catch (Exception exc)
                {
                    //si ca ne fonctionne pas, on ecrit ca:
                    System.Console.WriteLine("Ecriture de le log impossible : {0}", FILELOG);
                    System.Console.WriteLine("Exception : {0}", exc.Message);
                }
            }
        }  //fin de private static void Writelog(string FILELOG, string MESSAGE_LOG, string TYPE_MESSAGE) 
    } //fin de public class EngineeringConfigurationScript : EngineeringConfigurationScriptBase





    public class ElemRefPart
    {
        private int element_id;
        private string element_type;
        private string pres_ext_identity;
        private string pres_prim_pref;
        private int pres_item_type;
        private string pres_iap_iref;
        private int pres_selectable_element;
        private int pres_updating_type;
        private int pres_nominal;
        private int pres_suppl_pres_form;
        private int dba_method;





        // Constructeur
        public ElemRefPart()
        {

        }

        public int Element_id
        {
            get { return this.element_id; }
            set { this.element_id = value; }
        }

        public string Element_type
        {
            get { return this.element_type; }
            set { this.element_type = value; }
        }

        public string Pres_ext_identity
        {
            get { return this.pres_ext_identity; }
            set { this.pres_ext_identity = value; }
        }

        public string Pres_prim_pref
        {
            get { return this.pres_prim_pref; }
            set { this.pres_prim_pref = value; }
        }


        public int Pres_item_type
        {
            get { return this.pres_item_type; }
            set { this.pres_item_type = value; }
        }


        public string Pres_iap_iref
        {
            get { return this.pres_iap_iref; }
            set { this.pres_iap_iref = value; }
        }


        public int Pres_selectable_element
        {
            get { return this.pres_selectable_element; }
            set { this.pres_selectable_element = value; }
        }


        public int Pres_updating_type
        {
            get { return this.pres_updating_type; }
            set { this.pres_updating_type = value; }
        }


        public int Pres_nominal
        {
            get { return this.pres_nominal; }
            set { this.pres_nominal = value; }
        }


        public int Pres_suppl_pres_form
        {
            get { return this.pres_suppl_pres_form; }
            set { this.pres_suppl_pres_form = value; }
        }


        public int Dba_method
        {
            get { return this.dba_method; }
            set { this.dba_method = value; }
        }
    }

}//fin de namespace Nmx.DataEngineering.EngineeringConfigurationScript
