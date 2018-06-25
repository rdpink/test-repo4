using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows;
using System.Windows.Browser;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Threading;
using SL2C_Client.CServer;
using SL2C_Client.Init;
using SL2C_Client.Interfaces;
using SL2C_Client.ModelViewBinding;
using SL2C_Client.Statics;
using System.Windows.Media;

// TS 21.11.12 errors handled
namespace SL2C_Client.Model
{
    public class Processing : IProcessing
    {
        private CSEnumProfileWorkspace _workspace = CSEnumProfileWorkspace.workspace_default;

        public enum QueryMode
        {
            QueryPrev = -1,
            QueryNew = 0,
            QueryNext = 1,
            QueryFulltextAll = 2,               // die alte logik: es wird über alle volltextfelder verODERt
            QueryFulltextDetail = 3,            // noch nicht implementierte logik: einzelne Felder können abgefragt werden
            QuerySearchEngine = 4,              // neue logik: als suchstring wird aus dem feld "cmis:folder:suchstring" geliefert und in Processing entsprechend behandelt

            //QueryLast = 5,
            QueryRefresh = 5,                   // refresh (with filters only: TS 20.03.14 andere logik)

            QueryOther = 6,
            QueryMore = 7,                      // weitere daten einfach dazulesen (wie querynext aber ohne den cache vorher zu leeren)
        }

        // sortierungen aus der workspace definition des profils
        private List<CSOrderToken> _workspaceordertokens = new List<CSOrderToken>();

        // TS 01.10.13 sortierungen bei der suche, werden z.b. von listview aus dessen profil hier gesetzt
        private List<CSOrderToken> _listordertokens = new List<CSOrderToken>();

        // TS 01.10.13 sortierungen zum lesen der kinder, werden z.b. von treeview aus dessen profil hier gesetzt
        private List<CSOrderToken> _treeordertokens = new List<CSOrderToken>();

        // TS 20.01.16
        private List<string> _listdisplayproperties = new List<string>();

        private List<string> _treedisplayproperties = new List<string>();
        private CSEnumOptionPresets_listdisplaymode _listdisplaymode = CSEnumOptionPresets_listdisplaymode.listdisplaymode_columns;

        // TS 14.03.14 wird verwendet um nach QueryStatics auf einen bestimmten Ordner zu positionieren (z.B. für Posteingang oder Maileingang)
        // muss explizit von aussen gesetzt werden (z.B. in EDesktopProcessing)
        private string _defaultstaticfolderid = "";

        // TS 24.03.14 wenn gesetzt werden die suchbedingungen nur beim ersten mal in datcache_objects gespeichert und dann bei refresh immer verwendet
        // verwendet z.b. bei edp_gesamtsuche
        private bool _hasfixedquery = false;

        // Used for triggering unload/load all feature
        private static bool _treeallloadtriggered = false;

        // Used for temporarily unbinding the listview
        private static bool _unbindListView = false;

        // TS 01.04.15
        private static CSEnumProfileWorkspace _lastselectedworkspace = CSEnumProfileWorkspace.workspace_default;
        private static CSRootNode _lastselectedrootnode = null;

        // TS 17.03.17
        private ExtendedQueryValues _queryext_musthavevalues;
        private ExtendedQueryValues _queryext_shouldhavevalues;
        private ExtendedQueryValues _queryext_mustnothavevalues;



        #region constructors**************************************************************************************************************************************

        private static int _instancecounter = 0;

        public Processing(CSEnumProfileWorkspace workspace)
        {
            _workspace = workspace;
            // TS 01.10.13 workspace orders initialisieren
            _InitWorkspaceQueryOrders();

            if (Config.Config.debug_show_instancecount)
            {
                System.Threading.Interlocked.Increment(ref _instancecounter);
                if (!Config.Config.debug_show_instancecount_unloadonly) System.Diagnostics.Debug.WriteLine("ProcessingCount + : " + _instancecounter);
            }
        }

        ~Processing()
        {
            if (Config.Config.debug_show_instancecount)
            {
                System.Threading.Interlocked.Decrement(ref _instancecounter);
                System.Diagnostics.Debug.WriteLine("ProcessingCount - : " + _instancecounter);
            }
        }

        #endregion constructors**************************************************************************************************************************************

        #region properties****************************************************************************************************************************************

        // TS 01.04.15
        public static CSEnumProfileWorkspace LastSelectedWorkspace
        {
            get { return _lastselectedworkspace; }
            private set { _lastselectedworkspace = value; }
        }

        public static CSRootNode LastSelectedRootNode
        {
            get { return _lastselectedrootnode; }
            private set { _lastselectedrootnode = value; }
        }

        public CSEnumProfileWorkspace Workspace { get { return _workspace; } }

        // TS 01.10.13
        public List<CSOrderToken> WorkspaceOrderTokens { get { return _workspaceordertokens; } }

        public List<CSOrderToken> ListOrderTokens { get { return _listordertokens; } }
        public List<CSOrderToken> TreeOrderTokens { get { return _treeordertokens; } }

        public static bool UnbindListView
        {
            get { return _unbindListView; }
            set { _unbindListView = value; }
        }

        // TS 20.01.16
        public List<string> ListDisplayProperties { get { return _listdisplayproperties; } }

        public List<string> TreeDisplayProperties { get { return _treedisplayproperties; } }
        public CSEnumOptionPresets_listdisplaymode ListDisplayMode { get { return _listdisplaymode; } set { _listdisplaymode = value; } }

        /// <summary>
        /// liefert die komplette sortierung von workspace und liste
        /// </summary>
        public string QueryOrder
        {
            get
            {
                string orderby = "";
                List<CServer.CSOrderToken> givenorderby = new List<CSOrderToken>();
                _Query_GetSortDataFilter(ref givenorderby, true);
                foreach (CServer.CSOrderToken orderbytoken in givenorderby)
                {
                    if (orderby.Length > 0)
                    {
                        // TS 01.10.13 muss komma sein statt blank !!
                        //orderby = orderby + " ";
                        orderby = orderby + ", ";
                    }
                    orderby = orderby + orderbytoken.propertyname + " " + orderbytoken.orderby;
                }
                return orderby;
            }
        }

        /// <summary>
        /// liefert die komplette sortierung von workspace und tree
        /// </summary>
        public string StructOrder
        {
            get
            {
                string orderby = "";
                List<CServer.CSOrderToken> givenorderby = new List<CSOrderToken>();
                _Query_GetSortDataFilter(ref givenorderby, false);
                foreach (CServer.CSOrderToken orderbytoken in givenorderby)
                {
                    if (orderby.Length > 0)
                    {
                        // TS 01.10.13 muss komma sein statt blank !!
                        //orderby = orderby + " ";
                        orderby = orderby + ", ";
                    }
                    orderby = orderby + orderbytoken.propertyname + " " + orderbytoken.orderby;
                }
                return orderby;
            }
        }

        /// <summary>
        /// prüft ob zum aktuell selektierten objekt aktuell die relation angezeigt wird
        /// </summary>
        public bool IsRelationShown
        {
            get
            {
                bool ret = false;
                // nur in den NICHT related workspaces suchen
                string ws = this.Workspace.ToString();
                if (!ws.EndsWith("_related"))
                {
                    ws = ws + "_related";
                    CSEnumProfileWorkspace wsrelated = EnumHelper.GetCSEnumFromValue<CSEnumProfileWorkspace>(ws);
                    if (DataAdapter.Instance.DataCache.WorkspacesUsed.Contains(wsrelated))
                    {
                        IDocumentOrFolder dummy = DataAdapter.Instance.DataCache.Objects(Workspace).Object_Selected;
                        if (dummy.hasRelationships)
                        {
                            // TS 27.02.14 schleife drum damit nicht bei jedem klick im related workspace die ansicht zuklappt
                            IDocumentOrFolder relatedfound = null;
                            foreach (IDocumentOrFolder relatedselected in DataAdapter.Instance.DataCache.Objects(wsrelated).ObjectList)
                            {
                                foreach (DocOrFolderRelationship rel in dummy.Relationships)
                                {
                                    if (relatedselected.objectId.Length > 0
                                        && !relatedselected.objectId.Equals(DataAdapter.Instance.DataCache.Objects(wsrelated).Root.objectId)
                                        &&
                                        (relatedselected.objectId.Equals(rel.target_objectid) || relatedselected.objectId.Equals(rel.source_objectid)))
                                    {
                                        relatedfound = relatedselected;
                                        ret = true;
                                        break;
                                    }
                                }
                                // TS 27.02.14 ausstieg wenn gefunden
                                if (ret == true)
                                    break;
                            }
                            // TS 18.08.16 noch prüfen ob das gefundene auch selektiert ist
                            if (ret)
                            {
                                IDocumentOrFolder realrelatedselected = DataAdapter.Instance.DataCache.Objects(wsrelated).Object_Selected;
                                if (!relatedfound.objectId.Equals(realrelatedselected.objectId))
                                {
                                    DataAdapter.Instance.Processing(wsrelated).SetSelectedObject(relatedfound.objectId);
                                }
                            }
                        }
                    }
                }
                return ret;
            }
        }

        public IDocumentOrFolder ShownRelatedObject
        {
            get
            {
                IDocumentOrFolder relatedselected = new DocumentOrFolder(CSEnumProfileWorkspace.workspace_undefined);
                string ws = this.Workspace.ToString();
                if (!ws.EndsWith("_related"))
                {
                    ws = ws + "_related";
                    CSEnumProfileWorkspace wsrelated = EnumHelper.GetCSEnumFromValue<CSEnumProfileWorkspace>(ws);
                    if (DataAdapter.Instance.DataCache.WorkspacesUsed.Contains(wsrelated))
                    {
                        relatedselected = DataAdapter.Instance.DataCache.Objects(wsrelated).Object_Selected;
                    }
                }
                return relatedselected;
            }
        }

        // TS 14.03.14 wird verwendet um nach QueryStatics auf einen bestimmten Ordner zu positionieren (z.B. für Posteingang oder Maileingang)
        // muss explizit von aussen gesetzt werden (z.B. in EDesktopProcessing)
        public string DefaultStaticFolderId { get { return _defaultstaticfolderid; } set { _defaultstaticfolderid = value; } }

        public bool HasFixedQuery { get { return _hasfixedquery; } set { _hasfixedquery = value; } }

        private bool isMasterApplication { get { return ClientCommunication.Communicator.Instance.isMaster; } }

        // TS 17.03.17
        public ExtendedQueryValues QueryMustHaveValues { get { return _queryext_musthavevalues; } }
        public ExtendedQueryValues QueryShouldHaveValues { get { return _queryext_shouldhavevalues; } }
        public ExtendedQueryValues QueryMustNotHaveValues { get { return _queryext_mustnothavevalues; } }


        #endregion properties****************************************************************************************************************************************

        // *******************************************************************************************************************************************************
        // COMMANDING
        // *******************************************************************************************************************************************************

        // ==================================================================

        // ==================================================================

        #region AcknowledgePostObjectsToECM

        public void AcknowledgePostObjectsToECM(List<string> objectids)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (DataAdapter.Instance.DataCache.Objects(this.Workspace).ObjectIds_PreparedPostToECM.Count > 1)
                {

                    List<string> objectidlist = new List<string>();

                    // The first entry is the postcontainer
                    string postcontainerID = DataAdapter.Instance.DataCache.Objects(this.Workspace).ObjectIds_PreparedPostToECM.First().Value;

                    // then extract the object id's and build a string
                    string idString = postcontainerID + ";" + String.Join(";", objectids);

                    // Add the semikolon-separated id-string to the id-list
                    objectidlist.Add(idString);

                    List <cmisObjectType> cmisobjectlist = null;

                    CSRCPushPullUIRouting routinginfo = new CSRCPushPullUIRouting();
                    CSEnumRCPushUIListeners?[] receiver = new CSEnumRCPushUIListeners?[1];
                    receiver[0] = CSEnumRCPushUIListeners.HTML_CLIENT_DESKTOP;
                    routinginfo.sendto_componentorvirtualids = receiver;

                    CallbackAction callback = new CallbackAction(AcknowledgePostObjectsToECM_Done, objectids);
                    DataAdapter.Instance.DataProvider.RCPush_UI(routinginfo, CSEnumRCPushCommands.posttoecm, objectidlist, cmisobjectlist, DataAdapter.Instance.DataCache.Rights.UserPrincipal, callback);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        private void AcknowledgePostObjectsToECM_Done(object objectids)
        {
            if (DataAdapter.Instance.DataCache.ResponseStatus.success)
            {
                List<string> objectIDList = (List<string>)objectids;

                try
                {
                    foreach (string objectid in objectIDList)
                    {
                        if (DataAdapter.Instance.DataCache.Objects(this.Workspace).ObjectIds_PreparedPostToECM.Count > 1 &&
                            DataAdapter.Instance.DataCache.Objects(this.Workspace).ObjectIds_PreparedPostToECM.ContainsKey(objectid))
                        {
                            DataAdapter.Instance.DataCache.Objects(this.Workspace).ObjectIds_PreparedPostToECM.Remove(objectid);
                            // wenn nur noch der eintrag auf den postcontainer vorhanden ist dann den auch löschen
                            if (DataAdapter.Instance.DataCache.Objects(this.Workspace).ObjectIds_PreparedPostToECM.Count == 1)
                            {
                                DataAdapter.Instance.DataCache.Objects(this.Workspace).ObjectIds_PreparedPostToECM.Clear();
                            }
                        }
                    }
                }
                catch (System.Exception ex) { if (ex.Message.Length > 0) { Log.Log.Debug(ex.Message); } }
            }
        }


        #endregion AcknowledgePostObjectsToECM

        // ==================================================================

        #region ApplyACL

        public void ApplyACL(IDocumentOrFolder obj, bool objectOnly, bool setCustomRights)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanApplyACL)
                {
                    CallbackAction finalcallback = new CallbackAction(ApplyACLCallbackGetObjectsUpDown, obj);
                    DataAdapter.Instance.DataProvider.ApplyACL(obj.CMISObject, objectOnly, setCustomRights, DataAdapter.Instance.DataCache.Rights.UserPrincipal, finalcallback);

                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        public bool CanApplyACL { get { return _CanApplyACL(DataAdapter.Instance.DataCache.Objects(Workspace).Object_Selected); } }

        private bool _CanApplyACL()
        {
            return _CanApplyACL(DataAdapter.Instance.DataCache.Objects(Workspace).Object_Selected);
        }

        private bool _CanApplyACL(IDocumentOrFolder obj)
        {
            return DataAdapter.Instance.DataCache.ApplicationFullyInit;
        }

        private void ApplyACLCallbackGetObjectsUpDown(object obj)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                IDocumentOrFolder castObject = (IDocumentOrFolder)obj;
                castObject.ChildrenSkip = 0;
                this.GetObjectsUpDown(castObject, true, null);
            }
            catch (Exception e) { Log.Log.Error(e); if (e.Message.Length > 0) throw new Exception("", e); if (e.Message.Length == 0) throw e; }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        #endregion ApplyACL

        // ==================================================================

        #region BPM_AddObject (cmd_BPM_AddObject_Notification and cmd_BPM_AddObject_Withdrawing)

        public DelegateCommand cmd_BPM_AddObject_Notification { get { return new DelegateCommand(BPM_AddObject_Notification, _CanBPMAddObject); } }
        public DelegateCommand cmd_BPM_AddObject_Withdrawing { get { return new DelegateCommand(BPM_AddObject_Withdrawing, _CanBPMAddObject); } }

        public bool CanBPMAddObject { get { return _CanBPMAddObject(); } }

        private bool _CanBPMAddObject()
        {
            bool ret = false;
            try
            {
                string curriFrameObj = HtmlPage.Document.GetElementById("bpmSideDropFrame").GetAttribute("currentWF") != null ? HtmlPage.Document.GetElementById("bpmSideDropFrame").GetAttribute("currentWF").ToString() : "";
                if(curriFrameObj != null && curriFrameObj.Length > 0)
                {
                    bool isContaining = false;
                    IDocumentOrFolder objectToAdd = DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_default).Object_Selected;
                    IDocumentOrFolder workflowObject = DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_workflow).Folder_Selected_Level(Statics.Constants.WORKFLOW_TYPE_LEVEL);
                    if (!objectToAdd.isDocument)
                    {
                        foreach (IDocumentOrFolder childObj in workflowObject.ChildObjects)
                        {
                            if (childObj.objectId.Equals(objectToAdd.objectId))
                            {
                                isContaining = true;
                                break;
                            }
                        }

                        ret = !isContaining;
                    }
                }
                
            }
            catch (Exception e) { Log.Log.Error(e); }
            return ret;
        }

        /// <summary>
        /// Handler method for adding notification-objects to the workflow
        /// </summary>
        public void BPM_AddObject_Notification()
        {
            BPM_AddObject(Statics.Constants.WORKFLOW_OBJECT_NOTIFY);
        }

        /// <summary>
        /// Handler method for adding withdrawing-objects to the workflow
        /// </summary>
        public void BPM_AddObject_Withdrawing()
        {
            BPM_AddObject(Statics.Constants.WORKFLOW_OBJECT_NORMAL);
        }

        /// <summary>
        /// Adds selected object from treeview to the bpm-workflow via webservice-call
        /// </summary>
        public void BPM_AddObject(string workflow_type)
        {
            if (CanBPMAddObject)
            {
                try
                {
                    //Get List of selected cmis-object
                    List<IDocumentOrFolder> selObjects = new List<IDocumentOrFolder>();
                    selObjects.Add(DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_default).Object_Selected);

                    //Get parent-workflow
                    String parentWF = DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_workflow).Folder_Selected_Level(Statics.Constants.WORKFLOW_TYPE_LEVEL)[CSEnumCmisProperties.BPM_WORKFLOWID];

                    //Call Webservice
                    CallbackAction callback = new CallbackAction(BPMAddObject_Succeed);
                    DataAdapter.Instance.DataProvider.BPM_AddDocuments(parentWF, selObjects, workflow_type, DataAdapter.Instance.DataCache.Rights.UserPrincipal, callback);
                }
                catch (Exception e) { Log.Log.Error(e); }
            }
        }

        /// <summary>
        /// On successfull object-adding, refreshes the containing objects
        /// </summary>
        public void BPMAddObject_Succeed()
        {
            //Refresh Containing documents
            try
            {
                if (DataAdapter.Instance.DataCache.ResponseStatus.success)
                {
                    string objID = DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_workflow).Folder_Selected_Level(Statics.Constants.WORKFLOW_TYPE_LEVEL)[CSEnumCmisProperties.BPM_WORKFLOWID];
                    //CallbackAction callback = new CallbackAction(BPM_ToggleContentMap);
                    CallbackAction callback = null;
                    DataAdapter.Instance.DataProvider.BPM_GetDocuments(objID, CSEnumProfileWorkspace.workspace_workflow, DataAdapter.Instance.DataCache.Rights.UserPrincipal, callback);
                    DataAdapter.Instance.InformObservers(CSEnumProfileWorkspace.workspace_default);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
        }

        // ------------------------------------------------------------------

        #endregion BPM_AddObject (cmd_BPM_AddObject_Notification and cmd_BPM_AddObject_Withdrawing)

        // ==================================================================

        #region BPM_RemoveObject (cmd_BPM_RemoveObject)

        public DelegateCommand cmd_BPM_RemoveObject { get { return new DelegateCommand(BPM_RemoveObject, _CanBPMRemoveObject); } }
        public bool CanBPMRemoveObject { get { return _CanBPMRemoveObject(); } }

        private bool _CanBPMRemoveObject()
        {
            bool ret = false;
            try
            {
                string workflowID = DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_workflow).Folder_Selected_Level(Statics.Constants.WORKFLOW_TYPE_LEVEL)[CSEnumCmisProperties.BPM_WORKFLOWID];

                //Checks if one of the selected workflow-elements is a child of the selected workflow
                List<IDocumentOrFolder> selList = DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_workflow).Objects_Selected;
                foreach (IDocumentOrFolder selObj in selList)
                {
                    if (selObj.structLevel > Statics.Constants.WORKFLOW_TYPE_LEVEL && selObj.parentId.Equals(workflowID))
                    {
                        ret = true;
                        break;
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            return ret;
        }

        /// <summary>
        /// Removes the selected object from the content map off the workflow
        /// </summary>
        public void BPM_RemoveObject()
        {
            if (CanBPMRemoveObject)
            {
                try
                {
                    string workflowID = DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_workflow).Folder_Selected_Level(Statics.Constants.WORKFLOW_TYPE_LEVEL)[CSEnumCmisProperties.BPM_WORKFLOWID];
                    List<IDocumentOrFolder> selChilds = new List<IDocumentOrFolder>();

                    //Get selected child-object
                    List<IDocumentOrFolder> selList = DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_workflow).Objects_Selected;
                    foreach (IDocumentOrFolder selObj in selList)
                    {
                        if (selObj.structLevel > Statics.Constants.WORKFLOW_TYPE_LEVEL && selObj.parentId.Equals(workflowID))
                        {
                            //Remove selected object from datacache
                            List<string> remList = new List<string>();
                            remList.Add(selObj.objectId);
                            DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_workflow).RemoveObjects(remList);
                            DataAdapter.Instance.InformObservers(CSEnumProfileWorkspace.workspace_default);

                            //Add object to the collection
                            selChilds.Add(selObj);
                            break;
                        }
                    }

                    //Send childlist to the webservice
                    if (selChilds.Count > 0)
                    {
                        DataAdapter.Instance.DataProvider.BPM_RemoveDocuments(workflowID, selChilds, DataAdapter.Instance.DataCache.Rights.UserPrincipal);
                    }
                }
                catch (Exception e) { Log.Log.Error(e); }
            }
        }

        /// <summary>
        /// On successfull object-removing, refreshes the containing objects
        /// </summary>
        public void BPMRemoveObject_Succeed()
        {
            try
            {
                //Refresh Containing documents
                string objID = DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_workflow).Folder_Selected_Level(Statics.Constants.WORKFLOW_TYPE_LEVEL)[CSEnumCmisProperties.BPM_WORKFLOWID];
                //CallbackAction callback = new CallbackAction(BPM_ToggleContentMap);
                CallbackAction callback = null;
                DataAdapter.Instance.DataProvider.BPM_GetDocuments(objID, CSEnumProfileWorkspace.workspace_workflow, DataAdapter.Instance.DataCache.Rights.UserPrincipal, callback);
            }
            catch (Exception e) { Log.Log.Error(e); }
        }

        // ------------------------------------------------------------------

        #endregion BPM_RemoveObject (cmd_BPM_RemoveObject)

        // ==================================================================

        #region BPM_StartWF (cmd_BPM_StartWF)

        public DelegateCommand cmd_BPM_StartWF { get { return new DelegateCommand(BPM_StartWF, _CanBPM_StartWF); } }

        // zwei einstiege: 1. ohne parameter ruft den dialog
        // 2. mit parametern aus dem dialog gerufen startet dann wirklich einen workflow
        private void BPM_StartWF()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanBPM_StartWF)
                {
                    if (DataAdapter.Instance.DataCache.Meta[Statics.Constants.REPOSITORY_WORKFLOW.ToString()] == null
                    || DataAdapter.Instance.DataCache.Meta[Statics.Constants.REPOSITORY_WORKFLOW.ToString()].Count == 0)
                    {
                        DataAdapter.Instance.DataProvider.BPM_GetStartableTypes(DataAdapter.Instance.DataCache.Rights.UserPrincipal, new CallbackAction(BPM_StartWF_SourcesRead));
                    }
                    else
                        BPM_StartWF_SourcesRead();
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        private void BPM_StartWF_SourcesRead()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (DataAdapter.Instance.DataCache.ResponseStatus.success)
                {
                    View.Dialogs.dlgBPM child = new View.Dialogs.dlgBPM(this.Workspace);
                    DialogHandler.Show_Dialog(child);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public void BPM_StartWF(string workflowtype, List<IDocumentOrFolder> objectstoadd, bool copydocuments, CallbackAction callback)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (_CanBPM_StartWF(workflowtype, objectstoadd, copydocuments))
                {
                    List<cmisObjectType> cmisobjects = new List<cmisObjectType>();
                    foreach (IDocumentOrFolder obj in objectstoadd)
                    {
                        cmisobjects.Add(obj.CMISObject);
                    }
                    if (cmisobjects.Count > 0)
                    {
                        DataAdapter.Instance.DataProvider.BPM_StartWorkflow(workflowtype, cmisobjects, copydocuments, CSEnumProfileWorkspace.workspace_workflow, DataAdapter.Instance.DataCache.Rights.UserPrincipal, callback);
                    }else
                    {
                        DataAdapter.Instance.DataProvider.BPM_StartWorkflow(workflowtype, false, CSEnumProfileWorkspace.workspace_workflow, DataAdapter.Instance.DataCache.Rights.UserPrincipal, callback);
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        public bool CanBPM_StartWF { get { return _CanBPM_StartWF(); } }

        // TS 11.03.16
        //private bool _CanBPM_StartWF() { return DataAdapter.Instance.DataCache.ApplicationFullyInit; }
        private bool _CanBPM_StartWF() { return DataAdapter.Instance.DataCache.ApplicationFullyInit && DataAdapter.Instance.DataCache.Profile.UserProfile.bpmavail; }

        // TS 11.03.16
        private bool _CanBPM_StartWF(string workflowtype)
        {
            return
                _CanBPM_StartWF()
                && DataAdapter.Instance.DataCache.Meta[Statics.Constants.REPOSITORY_WORKFLOW.ToString()] != null
                && DataAdapter.Instance.DataCache.Meta[Statics.Constants.REPOSITORY_WORKFLOW.ToString()].Count > 0
                && DataAdapter.Instance.DataCache.Meta[Statics.Constants.REPOSITORY_WORKFLOW.ToString()].Contains(workflowtype);
        }

        private bool _CanBPM_StartWF(string workflowtype, List<IDocumentOrFolder> objectstoadd, bool copydocuments)
        {
            // TS 11.03.16
            //bool ret = DataAdapter.Instance.DataCache.ApplicationFullyInit
            //    && DataAdapter.Instance.DataCache.Meta[Statics.Constants.REPOSITORY_WORKFLOW.ToString()] != null
            //    && DataAdapter.Instance.DataCache.Meta[Statics.Constants.REPOSITORY_WORKFLOW.ToString()].Count > 0
            //    && DataAdapter.Instance.DataCache.Meta[Statics.Constants.REPOSITORY_WORKFLOW.ToString()].Contains(workflowtype);
            bool ret = _CanBPM_StartWF(workflowtype);

            // TS 11.03.16
            // if (ret && objectstoadd.Count > 0 && !copydocuments)
            if (ret && objectstoadd != null && objectstoadd.Count > 0 && !copydocuments)
            {
                foreach (IDocumentOrFolder child in objectstoadd)
                {
                    if (ret) ret = !child.isNotCreated;
                    if (!ret)
                        break;
                }
            }
            return ret;
        }

        public void BPM_StartWF_Succeed()
        {
            BPM_StartWF_Succeed(null);
        }

        /// <summary>
        /// Callback-Method for creating new workflows.
        /// Catches the new workflow-id and shows it on the screen.
        /// </summary>
        public void BPM_StartWF_Succeed(CallbackAction callback)
        {
            if (Log.Log.IsDebugEnabled)
                Log.Log.MethodEnter();

            try
            {
                if (DataAdapter.Instance.DataCache.ResponseStatus.success)
                {
                    //Get the last object from the workflow-datacache
                    IDocumentOrFolder objReturnFlow = null;
                    List<IDocumentOrFolder> returnList = new List<IDocumentOrFolder>();
                    returnList = DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_workflow).ObjectList.ToList();
                    objReturnFlow = returnList.Last();

                    //Get new Workflow-ID
                    string newWorkflowID = objReturnFlow[CSEnumCmisProperties.BPM_WORKFLOWID];

                    //Get Containing documents
                    DataAdapter.Instance.DataProvider.BPM_GetDocuments(newWorkflowID, CSEnumProfileWorkspace.workspace_workflow, DataAdapter.Instance.DataCache.Rights.UserPrincipal);
                    HtmlPage.Document.GetElementById("bpmSideDropFrame").SetAttribute("currentWF", newWorkflowID);

                    //Get Small-Mode-URL
                    string smallURL = objReturnFlow.GetCMISPropertyAllValues(CSEnumCmisProperties.BPM_SMALLMODE_URL).First();

                    //Navigate to my workflows, select new workflow, show Workflow in Sidedrop
                    DataAdapter.Instance.NavigateUI(CSEnumProfileWorkspace.workspace_edesktop, CSEnumProfileWorkspace.workspace_workflow);
                    DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_workflow).SetSelectedObject(objReturnFlow.objectId);
                    BPM_DisplaySideFrame(null, null, smallURL);
                    DataAdapter.Instance.InformObservers();                    

                    if (callback != null) callback.Invoke();
                }
            }
            catch (Exception e) { Log.Log.Error(e); }

            if (Log.Log.IsDebugEnabled)
                Log.Log.MethodLeave();
        }

        #endregion BPM_StartWF (cmd_BPM_StartWF)

        // ==================================================================

        #region Selected Workflow changed (BPM_SelectedWFChanged)

        /// <summary>
        /// Triggered, when the selected workflow is changed in EDesktopProcessing.cs
        /// </summary>
        /// <param name="oldWFID"></param>
        /// <param name="newWFID"></param>
        public void BPM_SelectedWFChanged(string newWorkflow)
        {
            try
            {
                HtmlElement bpmiFrame = HtmlPage.Document.GetElementById("bpmSideDropFrame");
                string currentWorkflow = bpmiFrame.GetAttribute("currentWF");
                string isVisible = bpmiFrame.GetStyleAttribute("display");

                // Hide the iFrame, if visible
                // TODO: Right behaviour? Maype the content should switch to the new iframe..
                if (isVisible == null && currentWorkflow != newWorkflow)
                {
                    if (newWorkflow != null)
                    {
                        IDocumentOrFolder newWorkflowObj = DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_workflow).GetObjectById(newWorkflow);
                        if (newWorkflowObj[CSEnumCmisProperties.BPM_WORKFLOW] != null)
                        {
                            BPM_DisplaySideFrame(null, null, newWorkflowObj[CSEnumCmisProperties.BPM_SMALLMODE_URL].ToString());

                            // TS 11.03.16 kinder einlesen wenn noch keine vorhanden
                            //string objID = DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_workflow).Folder_Selected_Level(Statics.Constants.WORKFLOW_TYPE_LEVEL)[CSEnumCmisProperties.BPM_WORKFLOWID];
                            //if (newWorkflowObj.ChildObjects.Count() == 0)
                            //{
                                //string objID = newWorkflowObj[CSEnumCmisProperties.BPM_WORKFLOWID];
                                //CallbackAction callback = new CallbackAction(BPM_ToggleContentMap);
                                //DataAdapter.Instance.DataProvider.BPM_GetDocuments(objID, CSEnumProfileWorkspace.workspace_workflow, DataAdapter.Instance.DataCache.Rights.UserPrincipal, callback);

                            //}
                            //else
                            //{
                                //BPM_ToggleContentMap();
                            //}
                        }
                        else
                        {
                            BPM_HideSideFrame(null, null);
                        }
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
        }

        // ------------------------------------------------------------------

        #endregion Selected Workflow changed (BPM_SelectedWFChanged)

        // ==================================================================

        #region Toggle Content Map (BPM_ToggleContentMap)

        /// <summary>
        /// Toggles visibility of the bpm-contentmap-overlay
        /// </summary>
        public void BPM_ToggleContentMap()
        {
            try
            {
                if (DataAdapter.Instance.DataCache.ResponseStatus.success)
                {
                    CSProfileComponent contentMap = DataAdapter.Instance.DataCache.Profile.Profile_GetComponentFromLayout(Statics.Constants.PROFILE_COMPONENTID_UMLAUFMAPPE);
                    IDocumentOrFolder parentObj = DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_workflow).Folder_Selected_Level(Statics.Constants.WORKFLOW_TYPE_LEVEL);

                    if (parentObj != null)
                    {
                        // Show map
                        string curriFrameObj = HtmlPage.Document.GetElementById("bpmSideDropFrame").GetAttribute("currentWF") != null ? HtmlPage.Document.GetElementById("bpmSideDropFrame").GetAttribute("currentWF").ToString() : "";
                        DataAdapter.Instance.DataProvider.BPM_GetDocuments(curriFrameObj, CSEnumProfileWorkspace.workspace_workflow, DataAdapter.Instance.DataCache.Rights.UserPrincipal, null);
                        if (curriFrameObj.Equals(parentObj.objectId))
                        {
                            ViewManager.ShowComponent(contentMap, true);
                            DispatcherTimer timer = new DispatcherTimer();
                            timer.Interval = new TimeSpan(0, 0, 0, 0, 20); // 100 ticks delay for updating UI
                            timer.Tick += delegate (object s, EventArgs ea)
                            {
                                timer.Stop();
                                timer.Tick -= delegate (object s1, EventArgs ea1) { };                                
                                DataAdapter.Instance.InformObservers(CSEnumProfileWorkspace.workspace_workflow);
                                DataAdapter.Instance.InformObservers(CSEnumProfileWorkspace.workspace_default);
                                timer = null;
                            };
                            timer.Start();

                        }
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
        }

        // ------------------------------------------------------------------

        #endregion Toggle Content Map (BPM_ToggleContentMap)

        // ==================================================================

        #region Display Related Object (BPM_DisplayRelatedObject)

        /// <summary>
        /// Anzeige der zugehörigen objekte im default workspace
        /// </summary>
        public void BPM_DisplayRelatedObject(IDocumentOrFolder selectedobject)
        {
            try
            {
                // Rebuild so that bpm-objects are always gathered from the webserver, for always getting the latest updated object (due to possible right-changes etc.)
                if (isMasterApplication)
                {
                    if (!selectedobject.RepositoryId.Equals(DataAdapter.Instance.DataCache.Rights.ChoosenRepository.repositoryId))
                    {
                        CSRootNode rootnode = DataAdapter.Instance.DataCache.RootNodes.GetNodeById(selectedobject.RepositoryId);
                        AppManager.RestartObject = selectedobject;
                        AppManager.RestartObjectWorkspace = CSEnumProfileWorkspace.workspace_default;
                        AppManager.ChooseRootNode(rootnode, false);
                    }else
                    {
                        DataAdapter.Instance.Processing(CSEnumProfileWorkspace.workspace_default).QuerySingleObjectById(selectedobject.RepositoryId, selectedobject.objectId, false, true, CSEnumProfileWorkspace.workspace_default, null);
                        DataAdapter.Instance.NavigateUI(CSEnumProfileWorkspace.workspace_default, CSEnumProfileWorkspace.workspace_default);
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
        }

        // ------------------------------------------------------------------

        #endregion Display Related Object (BPM_DisplayRelatedObject)

        // ==================================================================

        #region Cancel (cmd_Cancel)

        public DelegateCommand cmd_Cancel { get { return new DelegateCommand(Cancel, _CanCancel); } }
        public DelegateCommand cmd_CancelAll { get { return new DelegateCommand(CancelAll, _CanCancelAll); } }

        public void Cancel_Hotkey() { Cancel(); }
        public void Cancel()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanCancel)
                {
                    List<IDocumentOrFolder> cancelobjects = new List<IDocumentOrFolder>();

                    IDocumentOrFolder obj = DataAdapter.Instance.DataCache.Objects(Workspace).Object_Selected;
                    if (obj.isNotCreated || obj.isEdited || obj.canCheckIn || obj.isCut || obj.MoveParentId.Length > 0)
                    {
                        //cancelobjects.Add(obj);
                        cancelobjects = CollectParentAndChildrenToCancel(obj);
                    }
                    if (cancelobjects.Count > 0)
                    {
                        Cancel(cancelobjects, null);
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        private List<IDocumentOrFolder> CollectParentAndChildrenToCancel(IDocumentOrFolder parent)
        {
            List<IDocumentOrFolder> cancelobjects = new List<IDocumentOrFolder>();
            if (parent.isNotCreated || parent.isEdited || parent.canCheckIn || parent.isCut || parent.MoveParentId.Length > 0)
            {
                cancelobjects.Add(parent);
                foreach (IDocumentOrFolder child in parent.ChildObjects)
                {
                    List<IDocumentOrFolder> cancelchildren = CollectParentAndChildrenToCancel(child);
                    cancelobjects.AddRange(cancelchildren);
                }
            }
            return cancelobjects;
        }

        public void CancelAll_Hotkey() { CancelAll(); }
        public void CancelAll()
        {
            CancelAll(null);
        }
        public void CancelAll(CallbackAction callback)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanCancelAll)
                {
                    List<IDocumentOrFolder> cancelobjects = new List<IDocumentOrFolder>();
                    foreach (IDocumentOrFolder obj in DataAdapter.Instance.DataCache.Objects(Workspace).ObjectList_InWork.View)
                    {
                        if (obj.isNotCreated || obj.isEdited || obj.canCheckIn || obj.isCut || obj.MoveParentId.Length > 0)
                            cancelobjects.Add(obj);
                    }
                    Cancel(cancelobjects, callback);
                    DataAdapter.Instance.DataCache.Info.Viewer_KillSlideshow = true;
                }else
                {
                    if (callback != null)
                    {
                        callback.Invoke();
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public void Cancel(List<IDocumentOrFolder> cancelobjects, CallbackAction finalcallback)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (_CanCancel(cancelobjects))
                {
                    List<cmisObjectType> cancelcmisobjects = new List<cmisObjectType>();
                    foreach (IDocumentOrFolder child in cancelobjects)
                    {
                        cancelcmisobjects.Add(child.CMISObject);
                    }
                    if (cancelcmisobjects.Count > 0)
                    {
                        CallbackAction callback = new CallbackAction(Cancel_Done, cancelobjects, null);
                        if (finalcallback != null)
                            callback = new CallbackAction(Cancel_Done, cancelobjects, finalcallback);

                        DataAdapter.Instance.DataProvider.Cancel(cancelcmisobjects, DataAdapter.Instance.DataCache.Rights.UserPrincipal, callback);
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        public bool CanCancel { get { return _CanCancel(); } }

        public bool CanCancelAll { get { return _CanCancelAll(); } }

        private bool _CanCancel()
        {
            bool ret = DataAdapter.Instance.DataCache.ApplicationFullyInit;
            if (ret)
            {
                IDocumentOrFolder obj = DataAdapter.Instance.DataCache.Objects(Workspace).Object_Selected;
                ret = obj.isNotCreated || obj.isEdited || obj.canCheckIn || obj.isCut || obj.MoveParentId.Length > 0;
            }
            return ret;
        }

        // TS 28.08.14
        private bool _CanCancel(List<IDocumentOrFolder> cancelobjects)
        {
            bool ret = DataAdapter.Instance.DataCache.ApplicationFullyInit;
            return ret;
        }

        // TS 13.08.14
        private bool _CanCancelAll() { return (DataAdapter.Instance.DataCache.ApplicationFullyInit && DataAdapter.Instance.DataCache.Objects(Workspace).AnyInWork); }

        private void Cancel_Done(object cancelledobjects, CallbackAction finalcallback)
        {
            if (DataAdapter.Instance.DataCache.ResponseStatus.success)
            {
                if (DataAdapter.Instance.DataCache.Objects(this.Workspace).ObjectList.Count() == 1)
                    this.ClearCache();
                else
                {
                    RefreshDZCollectionsAfterProcess(cancelledobjects, null);
                }
                DataAdapter.Instance.DataCache.Objects(Workspace).ObjectId_EditMode = "";
                DataAdapter.Instance.InformObservers(Workspace);

                // Check for ECMDesktop-Relations and send back the acknowledgement
                List<IDocumentOrFolder> cancelObjects = (List<IDocumentOrFolder>)cancelledobjects;
                List<string> post2ecmObjects = new List<string>();
                foreach (IDocumentOrFolder cobject in cancelObjects)
                {
                    if (DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_default).ObjectIds_PreparedPostToECM.Count > 1
                        && DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_default).ObjectIds_PreparedPostToECM.ContainsKey(cobject.objectId)
                        && !cobject.isCut)
                    {
                        post2ecmObjects.Add("2C_" + cobject.RepositoryId + "_" + cobject.objectId);
                    }
                }
                AcknowledgePostObjectsToECM(post2ecmObjects);

                if (finalcallback != null) finalcallback.Invoke();
            }
        }

        // ------------------------------------------------------------------

        #endregion Cancel (cmd_Cancel)

        // die Change Aufrufe (Application, Mandant, Repository) sind dafür gedacht, um ohne Benutzeroberfläche zu wechseln
        // mit Oberfläche wird ChooseAppReptry oder ChooseMandant verwendet

        // ==================================================================

        #region ChangeApplication

        public void ChangeApplication(string applicationid)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanChangeApplication)
                {
                    // TS 10.09.15 umbau auf rootnodes
                    //AppManager.ChooseAppReptryLayout(applicationid, "", "", false);
                    CSRootNode node = DataAdapter.Instance.DataCache.RootNodes.GetFirstNodeForApplication(applicationid);
                    if (node != null)
                        AppManager.ChooseRootNode(node, false);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        private void ChangeApplication(object parameter)
        {
            ChangeApplication((string)parameter);
        }

        // ------------------------------------------------------------------
        public bool CanChangeApplication { get { return _CanChangeApplication(); } }

        private bool _CanChangeApplication()
        {
            return DataAdapter.Instance.DataCache.ApplicationFullyInit && DataAdapter.Instance.DataCache.Objects(Workspace).AnyInWork == false;
        }

        #endregion ChangeApplication

        // ==================================================================

        #region ChangeMandant

        public void ChangeMandant(string mandantid)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanChangeMandant)
                {
                    AppManager.ChooseMandant(mandantid, true);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        private void ChangeMandant(object parameter)
        {
            ChangeMandant((string)parameter);
        }

        // ------------------------------------------------------------------
        public bool CanChangeMandant { get { return _CanChangeMandant(); } }

        private bool _CanChangeMandant()
        {
            return DataAdapter.Instance.DataCache.ApplicationFullyInit && DataAdapter.Instance.DataCache.Objects(Workspace).AnyInWork == false
                && DataAdapter.Instance.DataCache.Rights.UserRights.mandanten.Length > 1;
        }

        #endregion ChangeMandant

        // ==================================================================

        #region ChangeRepository (cmd_ChangeRepository)

        // TS 19.02.15 das gibt es als command schon länger nicht mehr
        //public DelegateCommand cmd_ChangeRepository { get { return new DelegateCommand(ChangeRepository, _CanChangeRepository); } }
        public void ChangeRepository(string repositoryid)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanChangeRepository)
                {
                    // TS 10.09.15 umbau auf rootnodes
                    //AppManager.ChooseAppReptryLayout(DataAdapter.Instance.DataCache.Rights.ChoosenApplication.id, repositoryid, "", false);
                    CSRootNode node = DataAdapter.Instance.DataCache.RootNodes.GetNodeById(repositoryid);
                    if (node != null)
                        AppManager.ChooseRootNode(node, false);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        private void ChangeRepository(object parameter)
        {
            ChangeRepository((string)parameter);
        }

        // ------------------------------------------------------------------
        public bool CanChangeRepository { get { return _CanChangeRepository(); } }

        private bool _CanChangeRepository()
        {
            return DataAdapter.Instance.DataCache.ApplicationFullyInit && DataAdapter.Instance.DataCache.Objects(Workspace).AnyInWork == false;
        }

        #endregion ChangeRepository (cmd_ChangeRepository)

        // ==================================================================
        // TS 17.02.16 abgrenzung zu ChooseRootNode: Change wechselt auf den im Cache als TextProperty vorgewählten während Choose einen Dialog forciert

        #region ChangeRootNode (cmd_ChangeRootNode)

        public DelegateCommand cmd_ChangeRootNode { get { return new DelegateCommand(ChangeRootNode, _CanChangeRootNode); } }

        public void ChangeRootNode()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanChangeRootNode)
                {
                    CSRootNode nodetoselect = DataAdapter.Instance.DataCache.RootNodes.GetNodeByName(DataAdapter.Instance.DataCache.RootNodes.Node_Selected_Name);
                    AppManager.ChooseRootNode(nodetoselect, false);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        public bool CanChangeRootNode { get { return _CanChangeRootNode(); } }

        private bool _CanChangeRootNode()
        {
            bool ret = DataAdapter.Instance.DataCache.ApplicationFullyInit;
            if (ret) ret = DataAdapter.Instance.DataCache.Objects(Workspace).AnyInWork == false;
            if (ret) ret = DataAdapter.Instance.DataCache.RootNodes.NodeList_Useable.Count > 1;
            if (ret) ret = !DataAdapter.Instance.DataCache.RootNodes.Node_Selected_Name.Equals(DataAdapter.Instance.DataCache.RootNodes.Node_Selected.nodename);
            if (ret) ret = DataAdapter.Instance.DataCache.RootNodes.GetNodeByName(DataAdapter.Instance.DataCache.RootNodes.Node_Selected_Name) != null;
            return ret;
        }

        #endregion ChangeRootNode (cmd_ChangeRootNode)

        // ==================================================================


        #region ChangePassword (cmd_ChangePassword)

        public DelegateCommand cmd_ChangePassword { get { return new DelegateCommand(ChangePassword, _CanChangePassword); } }

        public void ChangePassword()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanChangePassword)
                {
                    View.Dialogs.dlgChangePassword changeDialog = new View.Dialogs.dlgChangePassword(false, null);
                    DialogHandler.Show_Dialog(changeDialog);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        public bool CanChangePassword { get { return _CanChangePassword(); } }

        private bool _CanChangePassword()
        {
            bool ret = DataAdapter.Instance.DataCache.ApplicationFullyInit && DataAdapter.Instance.DataCache.Rights.UserPrincipal.canchangepassword;
            return ret;
        }

        #endregion ChangePassword (cmd_ChangePassword)

        // ==================================================================

        #region ChooseRootNode (cmd_ChooseRootNode)

        public DelegateCommand cmd_ChooseRootNode { get { return new DelegateCommand(ChooseRootNode, _CanChooseRootNode); } }

        public void ChooseRootNode()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanChooseRootNode)
                {
                    AppManager.ChooseRootNode(null, true);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        public bool CanChooseRootNode { get { return _CanChooseRootNode(); } }

        private bool _CanChooseRootNode()
        {
            bool ret = DataAdapter.Instance.DataCache.ApplicationFullyInit;
            if (ret) ret = DataAdapter.Instance.DataCache.Objects(Workspace).AnyInWork == false;
            if (ret) ret = DataAdapter.Instance.DataCache.RootNodes.NodeList_Useable.Count > 1;
            return ret;
        }

        #endregion ChooseRootNode (cmd_ChooseRootNode)

        // ==================================================================

        #region ChooseMandant (cmd_ChooseMandant)

        public DelegateCommand cmd_ChooseMandant { get { return new DelegateCommand(ChooseMandant, _CanChooseMandant); } }

        public void ChooseMandant()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanChooseMandant)
                {
                    AppManager.ChooseMandant("", true);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        public bool CanChooseMandant { get { return _CanChooseMandant(); } }

        private bool _CanChooseMandant()
        {
            return (DataAdapter.Instance.DataCache.ApplicationFullyInit && DataAdapter.Instance.DataCache.Objects(Workspace).AnyInWork == false
             && DataAdapter.Instance.DataCache.Rights.UserRights.mandanten.Length > 1);
        }

        #endregion ChooseMandant (cmd_ChooseMandant)

        // ==================================================================

        #region ClearCache (cmd_ClearCache)

        public DelegateCommand cmd_ClearCache { get { return new DelegateCommand(_ClearCachePreRun, _CanClearCache); } }

        /// <summary>
        /// Prerunner for Clearcache-Commands to clear the fulltext-overall-cache
        /// </summary>
        private void _ClearCachePreRun()
        {
            try
            {                
                if (this.Workspace == CSEnumProfileWorkspace.workspace_default && CanClearCache)
                {
                    DataAdapter.Instance.DataCache.Profile.Option_SetValueBoolean(CSEnumOptions.searchovermatchlist, false);
                }
            }
            catch (System.Exception ex) { if (ex.Message.Length > 0) { Log.Log.Debug(ex.Message); } }
            
            // Merged ClearCache & CancelAll on Command
            CallbackAction callback = new CallbackAction(ClearCache);
            CancelAll(callback);
        }

        public void ClearCache_Hotkey() { _ClearCachePreRun(); }

        // TS 12.01.15 clearfulltextqueryvalues dazu
        public void ClearCache() { ClearCache(0, true, null); }

        public void ClearCache(bool clearfulltextqueryvalues, CallbackAction finalcallback)
        {
            ClearCache(0, clearfulltextqueryvalues, finalcallback);            
        }

        public void ClearCache(int callbacklevel, bool clearfulltextqueryvalues, CallbackAction finalcallback)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            List<string> object_ids_to_remove = new List<string>();
            try
            {
                if (CanClearCache)
                {
                    // TS 11.03.14 validation
                    UnbindListView = true;
                    DataAdapter.Instance.DataCache.Info.QueryIndexCount().Clear();
                    DataAdapter.Instance.DataCache.ClientCom.Update_ClientListeners(this.Workspace);
                    if (callbacklevel == 0)
                    {
                        bool ret = true;
                        bool autosave_create = DataAdapter.Instance.DataCache.Profile.Option_GetValueBoolean(CSEnumOptions.autosave_create);
                        bool autosave_edit = DataAdapter.Instance.DataCache.Profile.Option_GetValueBoolean(CSEnumOptions.autosave_edit);
                        List<IDocumentOrFolder> autosaves = new List<IDocumentOrFolder>();
                        List<IDocumentOrFolder> notautosaves = new List<IDocumentOrFolder>();
                        foreach (IDocumentOrFolder obj in DataAdapter.Instance.DataCache.Objects(this.Workspace).ObjectList_InWork.View)
                        {
                            // TS 28.05.14 fallunterscheidung
                            if ((obj.isNotCreated && autosave_create) || (obj.isEdited && autosave_edit))
                                autosaves.Add(obj);
                            else
                                notautosaves.Add(obj);
                        }

                        // automatisch speichern
                        if (autosaves.Count > 0 && notautosaves.Count == 0)
                        {
                            try
                            {
                                foreach (IDocumentOrFolder obj in autosaves)
                                {
                                    if (ret) ret = obj.ValidateData();
                                    if (!ret)
                                        break;
                                }
                            }
                            catch (Exception e) { Log.Log.Error(e); }
                            if (ret)
                            {
                                // TS 12.01.15 clearfulltextqueryvalues dazu
                                //this.Save(autosaves, new CallbackAction(ClearCache, 10, finalcallback));
                                this.Save(autosaves, new CallbackAction(ClearCache, 10, clearfulltextqueryvalues, finalcallback));
                            }
                            else
                            {
                                // meldung ausgeben
                                DisplayWarnMessage(LocalizationMapper.Instance["msg_warn_unsaved"]);
                                // das hier muß sein damit der button nicht einfach umschaltet (wird hiermit automatisch zurückgesetzt bzw. einfach refreshed)
                                DataAdapter.Instance.InformObservers(this.Workspace);
                            }
                        }
                        else if (notautosaves.Count > 0)
                        {
                            // meldung ausgeben
                            DisplayWarnMessage(LocalizationMapper.Instance["msg_warn_unsaved"]);
                            // das hier muß sein damit der button nicht einfach umschaltet (wird hiermit automatisch zurückgesetzt bzw. einfach refreshed)
                            DataAdapter.Instance.InformObservers(this.Workspace);
                        }
                        else
                            callbacklevel = 10;
                    }

                    if (callbacklevel == 10)
                    {
                        // TS 03.05.16
                        // DataAdapter.Instance.PauseOrResumeBindings(true);
                        DataAdapter.Instance.PauseOrResumeBindings(true, this.Workspace);

                        object_ids_to_remove.AddRange(DataAdapter.Instance.DataCache.Objects(Workspace).ClearCache(false));
                        DataAdapter.Instance.DataCache.Info.Viewer_KillSlideshow = true;

                        // TS 03.05.16
                        // DataAdapter.Instance.PauseOrResumeBindings(false);
                        DataAdapter.Instance.PauseOrResumeBindings(false, this.Workspace);

                        // TS 31.07.12 die beizubehaltenden daten sammeln und an den server mitgeben
                        List<cmisObjectType> remaining = new List<cmisObjectType>();
                        foreach (CSEnumProfileWorkspace filter in DataAdapter.Instance.DataCache.WorkspacesUsed)
                        {
                            List<IDocumentOrFolder> tmp = DataAdapter.Instance.DataCache.Objects(filter).GetAllDocuments();

                            // TS 20.02.13
                            // DIE REMAINING SIND ZU UMFANGREICH MIT ALL DEN PROPERTIES
                            // DER DATACACHE ERWARTET NUR:
                            //downloadname = p[0];
                            //dzfirstpage = p[1];
                            //dzallpages = p[2];
                            foreach (IDocumentOrFolder doc in tmp)
                            {
                                // TS 30.04.14 nicht die leeren mitgeben
                                if (!doc.isEmptyQueryObject && doc.objectId.Length > 0)
                                    remaining.Add(doc.CMISObject);
                            }
                        }

                        // TS 21.02.13 testweise rausgenommen
                        DataAdapter.Instance.DataCache.Objects(Workspace).CreateEmptyQueryObjects(DataAdapter.Instance.DataCache.Repository(Workspace).TypeDescendants);
                        // TS 03.12.13
                        //DataAdapter.Instance.InformObservers();
                        // TS 29.01.14 weiter nach unten
                        //DataAdapter.Instance.InformObservers(Workspace);

                        // externe anzeige leeren
                        // TS 20.02.13 rausgenommen weil "echt" extern angezeigt wird
                        //_DisplayExternal("", false, true);

                        // suchwerte zurücksetzen
                        // queryskip hier nicht zurücksetzen da auch bei QueryNext, QueryPrev der Cache geleert wird und damit die SkipInfo verloren wäre
                        // wird zurückgesetzt bei einer neuen Suche
                        //DataAdapter.Instance.DataCache.Objects(Workspace).QuerySkip = 0;
                        DataAdapter.Instance.DataCache.Objects(Workspace).QueryResultCount = 0;
                        DataAdapter.Instance.DataCache.Objects(Workspace).QueryHasMoreResults = false;

                        // TS 12.01.15
                        if (clearfulltextqueryvalues) DataAdapter.Instance.DataCache.Objects(Workspace).ClearFulltextSearchValues();

                        // TS 29.01.14 weiter nach unten
                        DataAdapter.Instance.InformObservers(Workspace);

                        // Clear Fulltext-Overall-Cache
                        if (Workspace == CSEnumProfileWorkspace.workspace_default)
                        {
                            object_ids_to_remove.AddRange(DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_searchoverall).ClearCache(false));
                            DataAdapter.Instance.InformObservers(CSEnumProfileWorkspace.workspace_searchoverall);
                        }

                        // auch auf dem server löschen (z.b. das tempdir)
                        // TS 02.08.13 callback dazu
                        //DataAdapter.Instance.DataProvider.ClearCache(remaining, DataAdapter.Instance.DataCache.Rights.UserPrincipal);
                        DataAdapter.Instance.DataProvider.ClearCache(object_ids_to_remove, remaining, DataAdapter.Instance.DataCache.Rights.UserPrincipal, finalcallback);
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            UnbindListView = false;
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        public bool CanClearCache { get { return _CanClearCache(); } }

        public bool _CanClearCache()
        {
            return (DataAdapter.Instance.DataCache.ApplicationFullyInit);
        }

        #endregion ClearCache (cmd_ClearCache)

        #region Clear_LocalStorage

        // ==================================================================
        public void Clear_LocalStorage_Hotkey()
        {
            Clear_LocalStorage();
        }

        public void Clear_LocalStorage()
        {
            if (showMessageBox(Localization.localstring.msg_ClearLocalstorage, MessageBoxButton.OKCancel) == MessageBoxResult.OK)
            {
                LocalStorage.Clear();
            }
        }

        #endregion
        // ==================================================================

        #region ClearClipboard (cmd_ClearClipboard)

        public DelegateCommand cmd_ClearClipboard { get { return new DelegateCommand(ClearClipboard, _CanClearClipboard); } }
        public DelegateCommand cmd_ClearClipboardSelected { get { return new DelegateCommand(ClearClipboardSelected, _CanClearClipboardSelected); } }

        public void ClearClipboard()
        {
            ClearClipboard(false);
        }

        public void ClearClipboardSelected()
        {
            ClearClipboard(true);
        }

        private void ClearClipboard(bool selectedonly)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                // Clear Clipboard On Listeners
                if (isMasterApplication)
                {
                    DataAdapter.Instance.DataCache.ClientCom.Update_ClientListeners_ClearClipboard();
                }

                // Clear local Clipboard
                if ((!selectedonly && CanClearClipboard) || (selectedonly && CanClearClipboardSelected))
                {
                    List<cmisObjectType> clearcmisobjects = new List<cmisObjectType>();

                    // TS 17.08.16
                    List<string> simpleremoveobjects = new List<string>();

                    if (!selectedonly)
                    {
                        foreach (IDocumentOrFolder obj in DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_clipboard).ObjectList)
                        {
                            if (obj.objectId != DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_clipboard).Root.objectId && obj.objectId.Length > 0)
                            {
                                // TS 17.08.16 nur wenn wir uns noch im gleichen archiv befinden, sonst einfach direkt die dinger im cache killen
                                //clearcmisobjects.Add(obj.CMISObject);
                                if (ClearClipboardCheckMustServerRemove(obj))
                                    clearcmisobjects.Add(obj.CMISObject);
                                else
                                    simpleremoveobjects.Add(obj.objectId);
                            }
                        }
                    }
                    else
                    {
                        foreach (IDocumentOrFolder obj in DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_clipboard).Objects_Selected)
                        {
                            if (obj.objectId != DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_clipboard).Root.objectId && obj.objectId.Length > 0)
                            {
                                // TS 17.08.16 nur wenn wir uns noch im gleichen archiv befinden, sonst einfach direkt die dinger im cache killen
                                //clearcmisobjects.Add(obj.CMISObject);
                                if (ClearClipboardCheckMustServerRemove(obj))
                                    clearcmisobjects.Add(obj.CMISObject);
                                else
                                    simpleremoveobjects.Add(obj.objectId);
                            }
                        }
                        // TS 29.08.14 wenn keins gefunden wurde, dann liegt das vermutlich daran,
                        // das nur eines vorhanden ist und dieses zwar farblich markiert aber nicht wirklich selektiert ist
                        // diesen fall einfach abprüfen und verarbeiten
                        if (clearcmisobjects.Count == 0 && DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_clipboard).Objects_Selected.Count == 1)
                        {
                            // TS 09.10.14 das klappt nicht wenn es ein dokument ist daher anders lösen
                            //if (DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_clipboard).Root.ChildObjects.Count > 0)
                            //{
                            //    clearcmisobjects.Add(DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_clipboard).Root.ChildObjects[0].CMISObject);
                            //}
                            if (DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_clipboard).ObjectList.Count == 2)
                            {
                                if (DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_clipboard).ObjectList[1].objectId !=
                                    DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_clipboard).Root.objectId)
                                {
                                    // TS 17.08.16 nur wenn wir uns noch im gleichen archiv befinden, sonst einfach direkt die dinger im cache killen
                                    //clearcmisobjects.Add(DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_clipboard).ObjectList[1].CMISObject);
                                    IDocumentOrFolder tmp = DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_clipboard).ObjectList[1];
                                    if (ClearClipboardCheckMustServerRemove(tmp))
                                        clearcmisobjects.Add(tmp.CMISObject);
                                    else
                                        simpleremoveobjects.Add(tmp.objectId);
                                }
                            }
                        }
                    }
                    // TS 17.08.16
                    //DataAdapter.Instance.DataProvider.ClearClipboard(clearcmisobjects, DataAdapter.Instance.DataCache.Rights.UserPrincipal, new CallbackAction(ClearClipboard_Done));
                    if (clearcmisobjects.Count > 0)
                        DataAdapter.Instance.DataProvider.ClearClipboard(clearcmisobjects, DataAdapter.Instance.DataCache.Rights.UserPrincipal, new CallbackAction(ClearClipboard_Done));
                    if (simpleremoveobjects.Count > 0)
                        DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_clipboard).RemoveObjects(simpleremoveobjects);
                }

                // Maybe Refresh The Listener-Clipboards
                if (isMasterApplication)
                {
                    Deployment.Current.Dispatcher.BeginInvoke((ThreadStart)(() =>
                    {
                        if (DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_clipboard).ObjectList.Count > 1)
                        {
                            foreach (IDocumentOrFolder clipboardObj in DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_clipboard).ObjectList)
                            {
                                DataAdapter.Instance.DataCache.ClientCom.Update_ClientAndMaster_OnCopy(clipboardObj);
                            }
                            // Set Selection
                            DataAdapter.Instance.DataCache.ClientCom.Update_ClientListeners(CSEnumProfileWorkspace.workspace_clipboard);
                        }
                    }));
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        private bool ClearClipboardCheckMustServerRemove(IDocumentOrFolder obj)
        {
            bool ret = true;
            string objectid = obj.objectId;
            string repid = obj.RepositoryId;
            CSEnumProfileWorkspace workspace = obj.Workspace;
            if (repid.Equals(DataAdapter.Instance.DataCache.Repository(workspace).RepositoryInfo.repositoryId) && DataAdapter.Instance.DataCache.Objects(workspace).ExistsObject(objectid))
                ret = true;
            else
                ret = false;
            return ret;
        }

        // ------------------------------------------------------------------
        public bool CanClearClipboard { get { return _CanClearClipboard(); } }

        public bool CanClearClipboardSelected { get { return _CanClearClipboardSelected(); } }

        private bool _CanClearClipboard()
        {
            // TS 08.11.13 nur wenn > 1 objekt in der liste ist (das erste ist nur die root)
            //return DataAdapter.Instance.DataCache.ApplicationFullyInit && DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_clipboard).ObjectList.Count > 0;
            return DataAdapter.Instance.DataCache.ApplicationFullyInit && DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_clipboard).ObjectList.Count > 1;
        }

        private bool _CanClearClipboardSelected()
        {
            return DataAdapter.Instance.DataCache.ApplicationFullyInit && DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_clipboard).Objects_Selected.Count > 0;
        }

        private void ClearClipboard_Done()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (DataAdapter.Instance.DataCache.ResponseStatus.success)
                {
                    //foreach (IDocumentOrFolder obj in DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_clipboard).ObjectList_Clip)
                    //{
                    //    CSEnumProfileWorkspace filter = obj.Workspace;
                    //    DataAdapter.Instance.DataCache.Objects(filter).ForceNotifyDescendants(obj);
                    //}

                    // TS 07.02.14 nur wenn leer
                    if (DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_clipboard).ObjectList.Count == 0)
                    {
                        DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_clipboard).ClearCache();
                        DataAdapter.Instance.InformObservers();
                        // TS 12.11.13 fenster schliessen
                        CSProfileComponent clip = DataAdapter.Instance.DataCache.Profile.Profile_GetFirstComponentOfType(CSEnumProfileComponentType.CLIPBOARD);
                        if (clip != null)
                            WindowClose(clip);
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        #endregion ClearClipboard (cmd_ClearClipboard)

        //cmd_ConvertDocuments
        // ==================================================================

        #region ConvertDocuments (cmd_ConvertDocuments)

        public DelegateCommand cmd_ConvertDocuments { get { return new DelegateCommand(ConvertDocuments, _CanConvertDocuments); } }

        private void ConvertDocuments()
        {
            // TS 13.07.17 dzColl_Optimization
            // DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected
            IDocumentOrFolder docselected = DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected;
            string optimized = docselected[CSEnumCmisProperties.dzColl_Optimization];
            if (optimized != null && optimized.Equals("1"))
            {
                ConvertDocument();
            }
            else
            {
                ConvertDocuments(DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected);
            }
        }

        private void ConvertDocuments(IDocumentOrFolder doctoconvert)
        {
            List<IDocumentOrFolder> list = new List<IDocumentOrFolder>();
            list.Add(doctoconvert);
            ConvertDocuments(list);
        }

        public void ConvertDocuments(List<IDocumentOrFolder> docstoconvert)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanConvertDocuments(docstoconvert))
                {
                    cmisObjectType parentfolder = null;
                    List<cmisObjectType> allcmisdocuments = new List<cmisObjectType>();
                    List<string> docidstoconvert = new List<string>();
                    IDocumentOrFolder parent = null;
                    foreach (IDocumentOrFolder obj in docstoconvert)
                    {
                        if (parentfolder == null)
                        {
                            parent = obj.ParentFolder;
                            parentfolder = obj.ParentFolder.CMISObject;
                        }
                        docidstoconvert.Add(obj.objectId);
                    }

                    _CollectDZImagesRecursive(parent, allcmisdocuments);
                    DataAdapter.Instance.DataProvider.ConvertDocuments(parentfolder, docidstoconvert, allcmisdocuments, true, DataAdapter.Instance.DataCache.Rights.UserPrincipal);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public void ConvertDocument()
        {
            IDocumentOrFolder doctoconvert = DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected;
            int pagefrom = 1;
            int pageto = doctoconvert.DZPageCountConverted;
            bool allpages = false;
            ConvertDocument(doctoconvert, pagefrom, pageto, allpages);
        }

        public void ConvertDocument(IDocumentOrFolder doctoconvert, int pagefrom, int pageto, bool allpages)
        {
            //int pagefrom = tmp.DZPageCountConverted;
            //string maxdownloads = DataAdapter.Instance.DataCache.Profile.Option_GetValue(CSEnumOptions.maxdownloads);
            //int pageto = pagefrom + int.Parse(maxdownloads);
            //if (pageto > tmp.DZPageCount)
            //    pageto = tmp.DZPageCount;
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (_CanConvertDocuments(doctoconvert))
                {
                    //cmisObjectType parentfolder = null;
                    List<cmisObjectType> allcmisdocuments = new List<cmisObjectType>();
                    List<string> docidstoconvert = new List<string>();
                    IDocumentOrFolder parent = null;

                    //cmisObjectType doc = doctoconvert.CMISObject;
                    IDocumentOrFolder tmpparent = DataAdapter.Instance.DataCache.Objects(Workspace).GetObjectById(doctoconvert.TempOrRealParentId);
                    cmisObjectType parentfolder = tmpparent.CMISObject;
                    parent = tmpparent;

                    docidstoconvert.Add(doctoconvert.objectId);
                    _CollectDZImagesRecursive(parent, allcmisdocuments);

                    // TS 01.02.16 auf nächste gelesene Seite positionieren
                    //DataAdapter.Instance.DataProvider.GetDocumentPages(doc, parent, cmisdocuments, pagefrom, pageto, DataAdapter.Instance.DataCache.Rights.UserPrincipal, new CallbackAction(GetDocumentPages, objectid, 1, finalcallback));

                    // erstmal kein callback
                    //DataAdapter.Instance.DataProvider.GetDocumentPages(doc, parent, cmisdocuments, pagefrom, pageto, DataAdapter.Instance.DataCache.Rights.UserPrincipal, new CallbackAction(GetDocumentPages, objectid, pagefrom, finalcallback));

                    DataAdapter.Instance.DataProvider.ConvertDocument(parentfolder, docidstoconvert, allcmisdocuments, pagefrom, pageto, allpages, DataAdapter.Instance.DataCache.Rights.UserPrincipal);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        private bool _CanConvertDocuments()
        {
            return _CanConvertDocuments(DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected);
        }

        private bool _CanConvertDocuments(IDocumentOrFolder obj)
        {
            List<IDocumentOrFolder> objlist = new List<IDocumentOrFolder>();
            objlist.Add(obj);
            return CanConvertDocuments(objlist);
        }

        public bool CanConvertDocuments(List<IDocumentOrFolder> objectlist)
        {
            bool canconvert = DataAdapter.Instance.DataCache.ApplicationFullyInit;
            if (canconvert)
            {
                foreach (IDocumentOrFolder doc in objectlist)
                {
                    // TS 08.02.16
                    //if (canconvert) canconvert = !doc.isDocumentTiled;
                    // TS 23.11.17 auch ungespeicherte mal testen
                    // if (canconvert) canconvert = doc.isDocument && !doc.isNotCreated && !doc.isDocumentTiled;
                    if (canconvert) canconvert = doc.isDocument && !doc.isDocumentTiled;
                }
            }
            return canconvert;
        }

        //private void ConvertDocumentsDone(string parentid)
        //{
        //    List<CServer.cmisObjectType> cmisdocuments = new List<CServer.cmisObjectType>();
        //    IDocumentOrFolder parent = DataAdapter.Instance.DataCache.Objects(this.Workspace).GetObjectById(parentid);
        //    if (parent.isFolder)
        //    {
        //        _CollectDZImagesRecursive(parent, cmisdocuments);
        //        DataAdapter.Instance.DataProvider.GetDZCollection(parent.CMISObject, cmisdocuments, DataAdapter.Instance.DataCache.Rights.UserPrincipal, null);
        //    }
        //}

        #endregion ConvertDocuments (cmd_ConvertDocuments)

        // ==================================================================

        #region Copy (cmd_Copy)

        public DelegateCommand cmd_Copy { get { return new DelegateCommand(Copy, _CanCopy); } }

        private void Copy()
        {
            Copy(DataAdapter.Instance.DataCache.Objects(Workspace).Objects_Selected);
        }

        public void Copy_Hotkey()
        {
            Copy(DataAdapter.Instance.DataCache.Objects(Workspace).Objects_Selected);
        }

        private void Copy(List<IDocumentOrFolder> objectstocopy)
        {
            Copy(objectstocopy, null);
        }

        /// <summary>
        /// wird verwendet als Befehl und aus Drag & Drop
        /// wenn kein neuer Parent mitgegeben dann wird in Clipboard kopiert (ohne Speichern)
        /// wenn neuer Parent mitgegeben dann wird direkt gespeichert => geändert TS 21.03.13: Speichern nur wenn Option autosave_drop gesetzt
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="target"></param>
        public void Copy(List<IDocumentOrFolder> objectstocopy, IDocumentOrFolder target)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();

            try
            {
                if (CanCopy(objectstocopy, target))
                {
                    List<cmisObjectType> copyobjects = new List<cmisObjectType>();
                    foreach (IDocumentOrFolder obj in objectstocopy)
                    {
                        copyobjects.Add(obj.CMISObject);
                    }

                    cmisObjectType cmistarget = null;

                    bool autosave = false;
                    bool copytoclip = true;
                    if (target != null)
                    {
                        copytoclip = false;
                        autosave = DataAdapter.Instance.DataCache.Profile.Option_GetValueBoolean(CSEnumOptions.autosave_drop);
                        cmistarget = target.CMISObject;
                    }
                    bool wasdocument = objectstocopy[objectstocopy.Count - 1].isDocument;
                    CSEnumProfileWorkspace ws = CSEnumProfileWorkspace.workspace_default;
                    if (copytoclip)
                        ws = CSEnumProfileWorkspace.workspace_clipboard;
                    else
                        ws = target.Workspace;

                    // Setup callbacks
                    CallbackAction callback = new CallbackAction(DataAdapter.Instance.Processing(ws).ObjectTransfer_Done, wasdocument);
                    CallbackAction finalCallback = null;
                    if(copytoclip)
                    {
                        finalCallback = new CallbackAction(CopyToClipboard_Done, objectstocopy, callback);
                    }else
                    {
                        finalCallback = callback;
                    }
                    DataAdapter.Instance.DataProvider.Copy(copyobjects, copytoclip, cmistarget, autosave, DataAdapter.Instance.DataCache.Rights.UserPrincipal, finalCallback);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        private void CopyToClipboard_Done(object copiedObjects, CallbackAction finalCallback)
        {
            // Send the copy-objects to the listening masters/listeners
            List<IDocumentOrFolder> copyObj = (List<IDocumentOrFolder>)copiedObjects;
            foreach(IDocumentOrFolder obj in copyObj)
            {
                DataAdapter.Instance.DataCache.ClientCom.Update_ClientAndMaster_OnCopy(obj);
            }

            // Clear Clipboard-Cache, if we are on a subpage
            if(!isMasterApplication)
            {
                DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_clipboard).ClearCache();
            }else
            // Sync Clipboard-Cache, if we are on the master-application
            {
                foreach(IDocumentOrFolder clipboardObj in DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_clipboard).ObjectList)
                {
                    DataAdapter.Instance.DataCache.ClientCom.Update_ClientAndMaster_OnCopy(clipboardObj);
                }
            }

            // Handle following callbacks
            if (finalCallback != null)
            {
                finalCallback.Invoke();
            }
        }

        private bool _CanCopy()
        {
            return _CanCopy(DataAdapter.Instance.DataCache.Objects(Workspace).Object_Selected, null);
        }

        private bool _CanCopy(IDocumentOrFolder obj, IDocumentOrFolder target)
        {
            List<IDocumentOrFolder> objlist = new List<IDocumentOrFolder>();
            objlist.Add(obj);
            return CanCopy(objlist, target);
        }

        public bool CanCopy(List<IDocumentOrFolder> objectlist, IDocumentOrFolder target)
        {
            bool cancopy = DataAdapter.Instance.DataCache.ApplicationFullyInit;
            foreach (IDocumentOrFolder obj in objectlist)
            {
                // TS 26.05.15 prüfung auf objektid dazu: && obj.objectId.Length > 0;
                if (cancopy) cancopy = !obj.objectId.Equals(DataAdapter.Instance.DataCache.Objects(this.Workspace).Root.objectId)
                                       && obj.RepositoryId.Length > 0
                                       && !obj.isClipboardObject
                                       && !DataAdapter.Instance.DataCache.Objects(Workspace).IsDescendantCut(obj.objectId)
                                       && !DataAdapter.Instance.DataCache.Objects(Workspace).HasDescendantCut(obj.objectId)
                                       && obj.objectId.Length > 0;
                if (cancopy && target != null)
                    cancopy = target.canCreateObjectLevel(obj.structLevel);
            }
            // TS 12.05.15
            //return cancopy;
            return cancopy && objectlist != null && objectlist.Count > 0;
        }

        #endregion Copy (cmd_Copy)

        // ==================================================================

        #region CreateDocImportClipboard (cmd_CreateDocImportClipboard)

        public DelegateCommand cmd_CreateDocImportClipboard { get { return new DelegateCommand(CreateDocImportClipboard, _CanCreateDocImportClipboard); } }

        public void CreateDocImportClipboard()
        {
            DataAdapter.Instance.Processing(CSEnumProfileWorkspace.workspace_clipboard).CreateDocImport(CSEnumDocumentCopyTypes.COPYTYPE_ORIGINAL, Constants.REFID_DEFAULT);
        }
        
        private bool _CanCreateDocImportClipboard()
        {
            return DataAdapter.Instance.Processing(CSEnumProfileWorkspace.workspace_clipboard).CanCreateDocImport(null, false);
        }

        #endregion CreateDocImportClipboard (cmd_CreateDocImport)

        // ==================================================================

        #region CreateDocImport (cmd_CreateDocImport)

        public DelegateCommand cmd_CreateDocImport { get { return new DelegateCommand(CreateDocImport, _CanCreateDocImport); } }

        public void CreateDocImport()
        {
            CreateDocImport(CSEnumDocumentCopyTypes.COPYTYPE_ORIGINAL, Constants.REFID_DEFAULT);
        }

        public void CreateDocImport(CSEnumDocumentCopyTypes copytype, string refid)
        {
            CreateDocImport(copytype, refid, null, null, null, false, null);
        }

        public void CreateDocImport(CSEnumDocumentCopyTypes copytype, string refid, IDocumentOrFolder parent, cmisTypeContainer doctypedefinition, cmisTypeContainer foldertypedefinition, bool setaddressflag, Action callback)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            string maxFileSize = "";
            string maxFileCount = "";
            try
            {
                // Clipboard-Import always without any parents
                if (this.Workspace == CSEnumProfileWorkspace.workspace_clipboard)
                {
                    parent = DataAdapter.Instance.DataCache.Objects(Workspace).Root;
                }else if(parent == null)
                {
                    // If no parent is given, use the selected one instead
                    parent = DataAdapter.Instance.DataCache.Objects(Workspace).Object_Selected;
                }

                // Get Max Filesize
                CSOption maxSizeLimitOption = DataAdapter.Instance.DataCache.Profile.Option_GetOption(CSEnumOptions.maxuploadsize);
                if (maxSizeLimitOption != null)
                {
                    maxFileSize = maxSizeLimitOption.value;
                }

                // Get Max Filecount
                CSOption maxCountLimitOption = DataAdapter.Instance.DataCache.Profile.Option_GetOption(CSEnumOptions.maxuploadcount);
                if (maxCountLimitOption != null)
                {
                    maxFileCount = maxCountLimitOption.value;
                }

                // New Logic: Let the LocalConnector handle the whole thing to get detailed file-information
                string lcexecutefile = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.LC_Command];
                string lcparamfilename = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.LC_Param];
                CSLCRequest lcrequest = CreateLCRequest();

                // Set Params 
                lcrequest.proc_parentobjectid = parent.objectId;
                lcrequest.proc_repositoryid = parent.RepositoryId;
                lcrequest.proc_param_doccopytype = copytype;
                lcrequest.proc_param_refid = refid;
                lcrequest.proc_param_max_file_size = maxFileSize;
                lcrequest.proc_param_max_file_count = maxFileCount;

                // Execute the LC-Call
                DataAdapter.Instance.DataProvider.LC_CreateCommandFile(CSEnumCommands.cmd_CreateDocImport, lcparamfilename, lcrequest, true, DataAdapter.Instance.DataCache.Rights.UserPrincipal);
                bool pushtolc = LC_IsPushEnabled;
                if (!pushtolc  && lcexecutefile != null && lcexecutefile.Length > 0 && lcparamfilename != null && lcparamfilename.Length > 0)
                {
                    LC_ExecuteCall(lcexecutefile, lcparamfilename);
                }

            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public void CreateDocImportContentPrepared(List<CServer.cmisContentStreamType> contents, CSEnumDocumentCopyTypes copytype, string refid, IDocumentOrFolder parent,
                                                    cmisTypeContainer doctypedefinition, cmisTypeContainer foldertypedefinition, bool setaddressflag, Action callback)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (contents.Count > 0)
                {
                    CreateDocuments(contents, copytype, refid, parent, doctypedefinition, foldertypedefinition, setaddressflag, callback);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        //public bool CanCreateDocImport { get { return _CanCreateDocImport(null, false); } }
        // TS 30.01.15
        public bool CanCreateDocImport(bool isaddress) { return CanCreateDocImport(null, isaddress); }

        private bool _CanCreateDocImport()
        {
            return CanCreateDocImport(null, false);
        }

        //private bool _CanCreateDocImport(IDocumentOrFolder parent)
        public bool CanCreateDocImport(IDocumentOrFolder parent, bool isaddress)
        {
            // TS 24.03.13 automatisch doklog anlegen wenn nicht vorhanden
            bool ret = false;
            if (LC_IsAvailable)
            {
                // Clipboard-Import always without any parents
                if (this.Workspace == CSEnumProfileWorkspace.workspace_clipboard)
                {
                    parent = DataAdapter.Instance.DataCache.Objects(Workspace).Root;
                    ret = true;
                }
                else
                {
                    parent = CreateObjectGetParent(parent);
                    if (parent.structLevel == Constants.STRUCTLEVEL_09_DOKLOG)
                        ret = parent.canCreateDocument && parent.canCreateObjectLevel(Constants.STRUCTLEVEL_10_DOKUMENT);
                    else
                    {
                        ret = parent.canCreateObjectLevel(Constants.STRUCTLEVEL_09_DOKLOG);
                        // TS 16.03.15 raus hier und generell prüfen
                        // ...
                    }
                    // TS 16.03.15 raus hier und generell prüfen
                    if (ret && isaddress)
                    {
                        ret = parent.isAddressObject;
                    }
                    else if (ret && !isaddress)
                    {
                        // TS 05.02.15
                        //ret = !parent.isAddressObject;
                        ret = parent.isPureOfficeObject;
                    }

                    if (parent.ACL != null)
                    {
                        bool readOnly = parent.hasCreatePermission(); //!parent.hasReadOnlyPermission();
                        ret = ret && readOnly;
                    }

                }
            }
            return ret;
        }

        #endregion CreateDocImport (cmd_CreateDocImport)

        // ==================================================================

        #region CreateDocIVersion (cmd_CreateDocIVersion)

        public DelegateCommand cmd_CreateDocIVersion { get { return new DelegateCommand(CreateDocIVersion, _CanCreateDocIVersion); } }

        public void CreateDocIVersion()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanCreateDocIVersion)
                {
                    CreateDocImport(CSEnumDocumentCopyTypes.COPYTYPE_CONTENT, DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected.objectId);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        public bool CanCreateDocIVersion { get { return _CanCreateDocIVersion(); } }

        private bool _CanCreateDocIVersion()
        {
            IDocumentOrFolder selObj = DataAdapter.Instance.DataCache.Objects(this.Workspace).Folder_Selected;

            return CanCreateDocuments
                && DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected.objectId.Length > 0
                && DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected.isNotCreated == false
                && DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected.canCheckIn == false
                && selObj.hasCreatePermission()
                &&
                (DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected.CopyType == CServer.CSEnumDocumentCopyTypes.COPYTYPE_ORIGINAL
                || DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected.CopyType == CServer.CSEnumDocumentCopyTypes.COPYTYPE_CONTENT);
        }

        #endregion CreateDocIVersion (cmd_CreateDocIVersion)

        // ==================================================================

        #region CreateDocTVersion (cmd_CreateDocTVersion)

        public DelegateCommand cmd_CreateDocTVersion { get { return new DelegateCommand(CreateDocTVersion, _CanCreateDocTVersion); } }

        public void CreateDocTVersion()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanCreateDocTVersion)
                {
                    CreateDocImport(CSEnumDocumentCopyTypes.COPYTYPE_TECHNICAL, DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected.objectId);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        public bool CanCreateDocTVersion { get { return _CanCreateDocTVersion(); } }

        private bool _CanCreateDocTVersion()
        {
            return CanCreateDocuments
                && DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected.objectId.Length > 0
                && DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected.isNotCreated == false
                && DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected.canCheckIn == false
                &&
                (DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected.CopyType == CServer.CSEnumDocumentCopyTypes.COPYTYPE_ORIGINAL
                || DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected.CopyType == CServer.CSEnumDocumentCopyTypes.COPYTYPE_CONTENT);
        }

        #endregion CreateDocTVersion (cmd_CreateDocTVersion)

        // ==================================================================
        // die beiden versionierungsfunktionen gibt es nun beim local connector
        // TS 16.04.14 wieder teilweise reingenommen (nicht als command) für richtext document in adressverwaltung

        #region CreateDocIVersionFrom (cmd_CreateDocIVersionFrom)

        //public DelegateCommand cmd_CreateDocIVersionFrom { get { return new DelegateCommand(CreateDocIVersionFrom, _CanCreateDocIVersionFrom); } }
        public void CreateDocIVersionFrom(IDocumentOrFolder sourcedoc, CallbackAction callback)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (_CanCreateDocIVersionFrom(sourcedoc))
                {
                    // TS 16.04.14
                    //IDocumentOrFolder dummy = DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected;
                    //DataAdapter.Instance.DataProvider.CreateDocumentFrom(dummy.CMISObject, CSEnumDocumentCopyTypes.COPYTYPE_CONTENT, DataAdapter.Instance.DataCache.Rights.UserPrincipal, new CallbackAction(CreateDocIVersionFrom_Done));
                    if (callback != null)
                        DataAdapter.Instance.DataProvider.CreateDocumentFrom(sourcedoc.CMISObject, CSEnumDocumentCopyTypes.COPYTYPE_CONTENT, DataAdapter.Instance.DataCache.Rights.UserPrincipal, callback);
                    else
                        DataAdapter.Instance.DataProvider.CreateDocumentFrom(sourcedoc.CMISObject, CSEnumDocumentCopyTypes.COPYTYPE_CONTENT, DataAdapter.Instance.DataCache.Rights.UserPrincipal);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        //public bool CanCreateDocIVersionFrom { get { return _CanCreateDocIVersionFrom(); } }
        private bool _CanCreateDocIVersionFrom(IDocumentOrFolder sourcedoc)
        {
            // TS 16.04.14
            //return CanCreateDocuments
            //    && DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected.canCheckOut
            //    && DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected.objectId.Length > 0
            //    && DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected.isNotCreated == false
            //    && DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected.canCheckIn == false;
            // TS 29.04.16
            //return CanCreateDocuments && sourcedoc.canCheckOut && sourcedoc.objectId.Length > 0 && sourcedoc.isNotCreated == false && sourcedoc.canCheckIn == false;
            return sourcedoc.canCheckOut && sourcedoc.objectId.Length > 0 && sourcedoc.isNotCreated == false && sourcedoc.canCheckIn == false;
        }

        // TS 16.12.13
        //private void CreateDocIVersionFrom_Done()
        //{
        //    if (DataAdapter.Instance.DataCache.ResponseStatus.success)
        //    {
        //        IDocumentOrFolder createddoc = DataAdapter.Instance.DataCache.Objects(Workspace).ObjectList[DataAdapter.Instance.DataCache.Objects(Workspace).ObjectList.Count - 1];
        //        if (createddoc.isDocument)
        //        {
        //            _RefreshDZCollectionsAfterProcess(createddoc);
        //            //IDocumentOrFolder parent = DataAdapter.Instance.DataCache.Objects(Workspace).GetObjectById(createddoc.parentId);
        //            //this._GetObjectsUpDown(parent, false, null);
        //        }
        //        this.SetSelectedObject(createddoc.objectId);

        //        // TS 17.12.13 wen automatisches speichern gewünscht war aber natürlich nicht erfolgt ist (weil dokument ja noch bearbeitet werden soll): meldung ausgeben
        //        bool autosave = DataAdapter.Instance.DataCache.Profile.Option_GetValueBoolean(CSEnumOptions.autosave_create);
        //        if (autosave)
        //            _DisplayWarnMessage(Localization.localstring.msg_autosave_document_notdoneversion);

        //        // Aufruf LC
        //        //LC_DocumentEdit(createddoc);
        //    }
        //}

        #endregion CreateDocIVersionFrom (cmd_CreateDocIVersionFrom)

        // ==================================================================
        // die beiden versionierungsfunktionen gibt es nun beim local connector

        #region CreateDocOVersionFrom (cmd_CreateDocOVersionFrom)

        //public DelegateCommand cmd_CreateDocOVersionFrom { get { return new DelegateCommand(CreateDocOVersionFrom, _CanCreateDocOVersionFrom); } }
        //public void CreateDocOVersionFrom()
        //{
        //    if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
        //    try
        //    {
        //        if (CanCreateDocOVersionFrom)
        //        {
        //            IDocumentOrFolder dummy = DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected;
        //            DataAdapter.Instance.DataProvider.CreateDocumentFrom(dummy.CMISObject, CSEnumDocumentCopyTypes.COPYTYPE_ORIGINAL, DataAdapter.Instance.DataCache.Rights.UserPrincipal, new CallbackAction(CreateDocOVersionFrom_Done));
        //        }
        //    }
        //    catch (Exception e) { Log.Log.Error(e); }
        //    if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        //}
        //// ------------------------------------------------------------------
        //public bool CanCreateDocOVersionFrom { get { return _CanCreateDocOVersionFrom(); } }
        //private bool _CanCreateDocOVersionFrom()
        //{
        //    return CanCreateDocuments
        //        && DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected.objectId.Length > 0
        //        && DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected.isNotCreated == false
        //        && DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected.canCheckIn == false;
        //}
        //private void CreateDocOVersionFrom_Done()
        //{
        //    if (DataAdapter.Instance.DataCache.ResponseStatus.success)
        //    {
        //        IDocumentOrFolder createddoc = DataAdapter.Instance.DataCache.Objects(Workspace).ObjectList[DataAdapter.Instance.DataCache.Objects(Workspace).ObjectList.Count - 1];
        //        if (createddoc.isDocument)
        //        {
        //            _RefreshDZCollectionsAfterProcess(createddoc);
        //            //IDocumentOrFolder parent = DataAdapter.Instance.DataCache.Objects(Workspace).GetObjectById(createddoc.parentId);
        //            //this._GetObjectsUpDown(parent, false, null);
        //        }
        //        this.SetSelectedObject(createddoc.objectId);

        //        // TS 17.12.13 wen automatisches speichern gewünscht war aber natürlich nicht erfolgt ist (weil dokument ja noch bearbeitet werden soll): meldung ausgeben
        //        bool autosave = DataAdapter.Instance.DataCache.Profile.Option_GetValueBoolean(CSEnumOptions.autosave_create);
        //        if (autosave)
        //            _DisplayWarnMessage(Localization.localstring.msg_autosave_document_notdoneversion);

        //        // Aufruf LC
        //        //LC_DocumentEdit(createddoc);
        //    }
        //}

        #endregion CreateDocOVersionFrom (cmd_CreateDocOVersionFrom)

        // ==================================================================

        #region CreateDocuments

        public void CreateDocuments(List<cmisContentStreamType> contents, CSEnumDocumentCopyTypes copytype, string refid)
        {
            CreateDocuments(contents, copytype, refid, null, null, null, false, null);
        }

        public void CreateDocuments(List<cmisContentStreamType> contents, CSEnumDocumentCopyTypes copytype, string refid, IDocumentOrFolder parent,
                                        cmisTypeContainer doctypedefinition, cmisTypeContainer foldertypedefinition, bool setaddressflag, Action callback)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (_CanCreateDocuments(parent, doctypedefinition))
                {
                    if (contents.Count > 0)
                    {
                        if (refid == null || refid.Length == 0) refid = Constants.REFID_DEFAULT;

                        // TS 23.06.15 für clipboard sonderbehandlung
                        if (this.Workspace == CSEnumProfileWorkspace.workspace_clipboard)
                        {
                            parent = DataAdapter.Instance.DataCache.Objects(Workspace).Root;
                            if (doctypedefinition == null)
                                doctypedefinition = DataAdapter.Instance.DataCache.Repository_FindTypeContainerForBaseId("cmisdocument");

                            DataAdapter.Instance.DataProvider.CreateDocuments(doctypedefinition, parent.CMISObject, contents, CSEnumDocumentCopyTypes.COPYTYPE_ORIGINAL, "", false, DataAdapter.Instance.DataCache.Rights.UserPrincipal, new CallbackAction(CreateDocuments_Done));
                        }
                        else
                        {
                            // parent holen
                            parent = CreateObjectGetParent(parent);                            
                            cmisObjectType cmisparent = parent.CMISObject;

                            // TS 26.03.14 wieder reingenommen mit prüfung auf pflichtfelder
                            bool autosave = DataAdapter.Instance.DataCache.Profile.Option_GetValueBoolean(CSEnumOptions.autosave_create);
                            if (parent.isNotCreated) autosave = false;

                            if (doctypedefinition == null)
                            {
                                doctypedefinition = DataAdapter.Instance.DataCache.Repository(Workspace).GetTypeContainerForStructLevel(Constants.STRUCTLEVEL_10_DOKUMENT);
                            }

                            // TS 21.12.17 mails anders verarbeiten
                            List<cmisContentStreamType> mailcontents = new List<cmisContentStreamType>();
                            List<cmisContentStreamType> othercontents = new List<cmisContentStreamType>();
                            foreach (cmisContentStreamType c in contents)
                            {
                                // TS 23.03.18
                                // if (c.filename.ToLower().EndsWith(".msg"))
                                if (c.filename.ToLower().EndsWith(Constants.DOCTYPE_MSG_EXTENSION))
                                {
                                    mailcontents.Add(c);
                                }
                                else
                                {
                                    othercontents.Add(c);
                                }
                            }

                            if (mailcontents.Count > 0)
                            {
                                // TS 21.12.17 mails jetzt immer automatisch speichern und als editiert zurueckliefern wie im aktenplan
                                // das macht insofern sinn, da diese z.b. per copy from outlook ueber den lc ebenfalls direkt gespeichert werden
                                bool autosavemails = true;

                                // ********************************************
                                // TS 23.03.18 msg dateien duerfen nicht direkt in doklogs angelegt werden da diese ja beim speichern in ein eigenes doklog gewandelt werden 
                                //if (callback != null)
                                //    DataAdapter.Instance.DataProvider.CreateDocuments(doctypedefinition, cmisparent, contents, copytype, refid, autosavemails, DataAdapter.Instance.DataCache.Rights.UserPrincipal, new CallbackAction(CreateDocuments_Done, new CallbackAction(callback)));
                                //else
                                //    DataAdapter.Instance.DataProvider.CreateDocuments(doctypedefinition, cmisparent, contents, copytype, refid, autosavemails, DataAdapter.Instance.DataCache.Rights.UserPrincipal, new CallbackAction(CreateDocuments_Done));

                                cmisObjectType mailcmisparent = parent.CMISObject;
                                bool stopmails = false;
                                if (parent.structLevel == Constants.STRUCTLEVEL_09_DOKLOG)
                                {
                                    if (parent.ParentFolder != null && parent.ParentFolder.structLevel > 0 && parent.ParentFolder.canCreateObjectLevel(Constants.STRUCTLEVEL_09_DOKLOG))
                                    {
                                        mailcmisparent = parent.ParentFolder.CMISObject;
                                        //DisplayMessage("MSG Dateien können nicht in ein bestehendes Dokument eingefügt werden. Sie werden nun automatisch darüber abgelegt.");
                                    }
                                    else if (parent.ParentFolder == null || parent.ParentFolder.structLevel == 0 || !parent.ParentFolder.canCreateObjectLevel(Constants.STRUCTLEVEL_09_DOKLOG))
                                    {
                                        stopmails = true;
                                        //showOKDialog("MSG Dateien können nicht in ein bestehendes Dokument eingefügt werden.");
                                        showOKDialog(LocalizationMapper.Instance["msg_mailimport_invalidparent"]);
                                        //DisplayMessage("MSG Dateien können nicht in ein bestehendes Dokument eingefügt werden und es konnte kein anderer Ablageort ermittelt werden.");
                                    }
                                }
                                if (!stopmails)
                                {
                                    if (callback != null)
                                        DataAdapter.Instance.DataProvider.CreateDocuments(doctypedefinition, mailcmisparent, contents, copytype, refid, autosavemails, DataAdapter.Instance.DataCache.Rights.UserPrincipal, new CallbackAction(CreateDocuments_Done, new CallbackAction(callback)));
                                    else
                                        DataAdapter.Instance.DataProvider.CreateDocuments(doctypedefinition, mailcmisparent, contents, copytype, refid, autosavemails, DataAdapter.Instance.DataCache.Rights.UserPrincipal, new CallbackAction(CreateDocuments_Done));
                                }
                                // ********************************************
                            }

                            // und nun die anderen dokumente verarbeiten
                            if (othercontents.Count > 0)
                            {
                                if (parent.structLevel != Constants.STRUCTLEVEL_09_DOKLOG && !parent.canCreateDocument)
                                {
                                    if (foldertypedefinition == null)
                                    {
                                        foldertypedefinition = DataAdapter.Instance.DataCache.Repository(Workspace).GetTypeContainerForStructLevel(Constants.STRUCTLEVEL_09_DOKLOG, parent);
                                        foldertypedefinition = AddDefaultValuesToTypeContainer(foldertypedefinition, setaddressflag);
                                    }

                                    // TS 18.06.15 autosave generell umgebaut
                                    autosave = false;
                                    // TS 21.12.17 das macht offensichtlich keinen sinn hier nachdem es oben auf false gesetzt wurde
                                    //if (autosave) autosave = CreateFolderCheckCanAutosave(foldertypedefinition, setaddressflag);

                                    // TS 30.01.15
                                    //if (CanCreateFolder(Constants.STRUCTLEVEL_09_DOKLOG, parent, foldertypedefinition))
                                    if (CanCreateFolder(Constants.STRUCTLEVEL_09_DOKLOG, parent, foldertypedefinition, setaddressflag))
                                    {
                                        // TS 23.04.15
                                        // TS 21.12.17 das macht offensichtlich keinen sinn hier nachdem es oben auf false gesetzt wurde
                                        //if (autosave) foldertypedefinition = AddDefaultValuesToTypeContainer(foldertypedefinition, setaddressflag);

                                        if (callback != null)
                                            DataAdapter.Instance.DataProvider.CreateDocumentsInFolders(doctypedefinition, foldertypedefinition, cmisparent, othercontents, setaddressflag, autosave, DataAdapter.Instance.DataCache.Rights.UserPrincipal, new CallbackAction(CreateDocuments_Done, new CallbackAction(callback)));
                                        else
                                            DataAdapter.Instance.DataProvider.CreateDocumentsInFolders(doctypedefinition, foldertypedefinition, cmisparent, othercontents, setaddressflag, autosave, DataAdapter.Instance.DataCache.Rights.UserPrincipal, new CallbackAction(CreateDocuments_Done));
                                    }
                                }
                                else
                                {
                                    if (callback != null)
                                        DataAdapter.Instance.DataProvider.CreateDocuments(doctypedefinition, cmisparent, othercontents, copytype, refid, autosave, DataAdapter.Instance.DataCache.Rights.UserPrincipal, new CallbackAction(CreateDocuments_Done, new CallbackAction(callback)));
                                    else
                                        DataAdapter.Instance.DataProvider.CreateDocuments(doctypedefinition, cmisparent, othercontents, copytype, refid, autosave, DataAdapter.Instance.DataCache.Rights.UserPrincipal, new CallbackAction(CreateDocuments_Done));
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        public bool CanCreateDocuments { get { return _CanCreateDocuments(null, null); } }

        //private bool _CanCreateDocuments()
        //{
        //    bool ret = false;
        //    IDocumentOrFolder dummy = DataAdapter.Instance.DataCache.Objects(Workspace).Folder_Selected;
        //    // CSEnumObjectStructLevel.DOKLOG
        //    if (dummy.structLevel == Constants.STRUCTLEVEL_09_DOKLOG)
        //        ret = dummy.canCreateDocument && dummy.canCreateObjectLevel(10);
        //    else
        //        ret = dummy.canCreateObjectLevel(Constants.STRUCTLEVEL_09_DOKLOG);
        //    return ret;
        //}
        private bool _CanCreateDocuments(IDocumentOrFolder parent, cmisTypeContainer typedef)
        {
            bool cancreate = false;
            // Clipboard-Import always without any parents
            if (this.Workspace == CSEnumProfileWorkspace.workspace_clipboard)
            {
                parent = DataAdapter.Instance.DataCache.Objects(Workspace).Root;
                cancreate = true;
            }
            else
            {
                parent = CreateObjectGetParent(parent);
                if (parent.structLevel == Constants.STRUCTLEVEL_09_DOKLOG)
                {
                    if (typedef == null)
                        cancreate = parent.canCreateDocument && parent.canCreateObjectLevel(Constants.STRUCTLEVEL_10_DOKUMENT);
                    else
                        cancreate = parent.canCreateDocument && parent.canCreateObjectTypeId(typedef.type.id);
                }
                else
                    cancreate = parent.canCreateObjectLevel(Constants.STRUCTLEVEL_09_DOKLOG);
            }
            return cancreate;
        }

        private void CreateDocuments_Done()
        {
            CreateDocuments_Done(null);
        }

        private void CreateDocuments_Done(CallbackAction callback)
        {
            if (DataAdapter.Instance.DataCache.ResponseStatus.success)
            {
                IDocumentOrFolder createddoc = DataAdapter.Instance.DataCache.Objects(Workspace).ObjectList[DataAdapter.Instance.DataCache.Objects(Workspace).ObjectList.Count - 1];

                // TS 21.12.17 wenn auf einem anhang/schnipsel stehend dann das zugehörige original selektieren
                if (createddoc.RefId != null && createddoc.RefId.Length > 0 && !createddoc.RefId.Equals("0") && createddoc.CopyType != CSEnumDocumentCopyTypes.COPYTYPE_ORIGINAL && createddoc.CopyType != CSEnumDocumentCopyTypes.COPYTYPE_CONTENT)
                {
                    IDocumentOrFolder dummydoc = DataAdapter.Instance.DataCache.Objects(Workspace).GetObjectById(createddoc.RefId);
                    if (dummydoc.objectId.Equals(createddoc.RefId))
                    {
                        createddoc = dummydoc;
                    }                    
                }

                // TS 08.06.18 nicht selektioeren wenn da nichts gescheite zurueck kam
                if (!createddoc.isNotAvailable)
                {
                    IDocumentOrFolder createdFolder = null;
                    if (createddoc.isDocument)
                    {
                        DataAdapter.Instance.DataCache.Info.ListView_ForcePageSelection = true;
                        this.SetSelectedObject(createddoc.objectId);
                        createddoc = createddoc.ParentFolder;
                        createdFolder = createddoc.ParentFolder;
                    }
                    else
                    {
                        createdFolder = createddoc;
                    }
                    this.SetSelectedObject(createddoc.objectId);

                    // Update the Parent-Folder
                    if (createdFolder.structLevel == Constants.STRUCTLEVEL_09_DOKLOG)
                    {
                        createdFolder = createdFolder.ParentFolder;
                    }
                    if (!createdFolder.hasChildFolders)
                    {
                        createdFolder.SetPropertyValue(CSEnumCmisProperties.hasChildFolders.ToString(), "true", true);
                        createdFolder.SetPropertyValue(CSEnumCmisProperties.childrenHasMore.ToString(), "false", true);
                    }
                }
            }
            if (callback != null)
                callback.Invoke();
        }

        #endregion CreateDocuments

        // ==================================================================

        #region CreateFolder_Level1 (cmd_CreateFolder_Level1)

        public DelegateCommand cmd_CreateFolder_Level1 { get { return new DelegateCommand(CreateFolder_Level1, _CanCreateFolder_Level1); } }

        public void CreateFolder_Level1()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanCreateFolder_Level1)
                    CreateFolder(Constants.STRUCTLEVEL_01_LAND);
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        public bool CanCreateFolder_Level1 { get { return _CanCreateFolder_Level1(); } }

        private bool _CanCreateFolder_Level1()
        {
            return CanCreateFolder(Constants.STRUCTLEVEL_01_LAND);
        }

        // ------------------------------------------------------------------

        #endregion CreateFolder_Level1 (cmd_CreateFolder_Level1)

        // ==================================================================

        #region CreateFolder_Level2 (cmd_CreateFolder_Level2)

        public DelegateCommand cmd_CreateFolder_Level2 { get { return new DelegateCommand(CreateFolder_Level2, _CanCreateFolder_Level2); } }

        public void CreateFolder_Level2()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanCreateFolder_Level2)
                    CreateFolder(Constants.STRUCTLEVEL_02_ORT);
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        public bool CanCreateFolder_Level2 { get { return _CanCreateFolder_Level2(); } }

        private bool _CanCreateFolder_Level2()
        {
            return CanCreateFolder(Constants.STRUCTLEVEL_02_ORT);
        }

        // ------------------------------------------------------------------

        #endregion CreateFolder_Level2 (cmd_CreateFolder_Level2)

        // ==================================================================

        #region cmd_PageAllDocuments (cmd_PageAllDocuments)

        public DelegateCommand cmd_PageAllDocuments { get { return new DelegateCommand(PageAllDocuments, _CanPageAllDocuments); } }

        /// <summary>
        /// Fetches all the selected Object's descendant documents for paging their files
        /// </summary>
        public void PageAllDocuments()
        {
            IDocumentOrFolder parent = null;

            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                // In the clipboard, use the root as parent
                if (this.Workspace == CSEnumProfileWorkspace.workspace_clipboard)
                {
                    parent = DataAdapter.Instance.DataCache.Objects(Workspace).Root;
                }
                else if (parent == null)
                {
                    // If no parent is given, use the selected one instead
                    parent = DataAdapter.Instance.DataCache.Objects(Workspace).Object_Selected;
                }

                // Fetch the display-Properties
                List<string> displayProps = _Query_GetDisplayProperties(true, true, DataAdapter.Instance.Processing(CSEnumProfileWorkspace.workspace_searchoverall));
                
                // Fetch Query-Token for the current Repository
                Dictionary<string, List<CSQueryToken>> repositoryTokens = new Dictionary<string, List<CSQueryToken>>();
                List<CSQueryToken> qTokenList = new List<CSQueryToken>();

                CSQueryToken qTokenInTree = new CSQueryToken();
                qTokenInTree.propertyname = "IN_TREE";
                qTokenInTree.propertyvalue = parent.objectId;
                qTokenInTree.propertytype = enumPropertyType.@string;
                qTokenInTree.propertyreptypeid = "9_9";
                qTokenList.Add(qTokenInTree);
                repositoryTokens.Add(parent.RepositoryId, qTokenList);

                // Fetch the Sortorder-Tokens for the current Repository (only sort for Document-Date and Folge)
                Dictionary<string, List<CSOrderToken>> repoSort = new Dictionary<string, List<CSOrderToken>>();
                List<CSOrderToken> tokenList = new List<CSOrderToken>();
                CSOrderToken token = new CSOrderToken();
                token.propertyname = CSEnumCmisProperties.DOKUMENTSYS_04.ToString(); ;
                token.orderby = CSEnumOrderBy.desc;
                tokenList.Add(token);
                repoSort.Add(parent.RepositoryId, tokenList);

                // Create a Query for all the Objects descendant-Documents
                DataAdapter.Instance.Processing(CSEnumProfileWorkspace.workspace_searchoverall).ClearCache();
                Query(displayProps, repositoryTokens, repoSort, false, true, false, true, false, new CallbackAction(PageAllDocuments_QueryDone));

            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        /// <summary>
        /// When the Query succeeds, trigger the Slideshow-Mode
        /// </summary>
        private void PageAllDocuments_QueryDone()
        {
            if (DataAdapter.Instance.DataCache.ResponseStatus.success)
            {
                // Fetch the Document-IDs in the Cache
                List<string> documentList = new List<string>();
                foreach(IDocumentOrFolder docCache in DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_searchoverall).ObjectList)
                {
                    if(docCache.isFolder && docCache.structLevel == Constants.STRUCTLEVEL_09_DOKLOG && docCache.objectId.Length > 0)
                    {
                        documentList.Add(docCache.objectId);
                    }
                }

                // Update the Slideshow-Objectlist
                DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_searchoverall).ObjectIdList_SlideShow.Clear();
                DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_searchoverall).ObjectIdList_SlideShow.AddRange(documentList);

                // Start the Slideshow
                DataAdapter.Instance.DataCache.Info.Viewer_Slideshow_Workspace = CSEnumProfileWorkspace.workspace_searchoverall;
                DataAdapter.Instance.DataCache.Info.Viewer_SlideshowPaging = true;
                DataAdapter.Instance.DataCache.Info.Viewer_StartSlideshow = true;                
            }
        }

        private bool _CanPageAllDocuments()
        {
            IDocumentOrFolder parent = null;

            // In the clipboard, use the root as parent
            if (this.Workspace == CSEnumProfileWorkspace.workspace_clipboard)
            {
                parent = DataAdapter.Instance.DataCache.Objects(Workspace).Root;
            }
            else if (parent == null)
            {
                // If no parent is given, use the selected one instead
                parent = DataAdapter.Instance.DataCache.Objects(Workspace).Object_Selected;
            }

            return CanPageAllDocuments(parent);
        }

        public bool CanPageAllDocuments(IDocumentOrFolder parent)
        {
            
            return DataAdapter.Instance.DataCache.ApplicationFullyInit && parent != null && parent.hasChildObjects && parent.isFolder;
        }

        #endregion CreateDocImport (cmd_CreateDocImport)

        // ==================================================================

        #region CreateFolder_Level3 (cmd_CreateFolder_Level3)

        public DelegateCommand cmd_CreateFolder_Level3 { get { return new DelegateCommand(CreateFolder_Level3, _CanCreateFolder_Level3); } }

        public void CreateFolder_Level3()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanCreateFolder_Level3)
                    CreateFolder(Constants.STRUCTLEVEL_03_STRASSE);
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        public bool CanCreateFolder_Level3 { get { return _CanCreateFolder_Level3(); } }

        private bool _CanCreateFolder_Level3()
        {
            return CanCreateFolder(Constants.STRUCTLEVEL_03_STRASSE);
        }

        // ------------------------------------------------------------------

        #endregion CreateFolder_Level3 (cmd_CreateFolder_Level3)

        // ==================================================================

        #region CreateFolder_Level4 (cmd_CreateFolder_Level4)

        public DelegateCommand cmd_CreateFolder_Level4 { get { return new DelegateCommand(CreateFolder_Level4, _CanCreateFolder_Level4); } }

        public void CreateFolder_Level4()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanCreateFolder_Level4)
                    CreateFolder(Constants.STRUCTLEVEL_04_HAUS);
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        public bool CanCreateFolder_Level4 { get { return _CanCreateFolder_Level4(); } }

        private bool _CanCreateFolder_Level4()
        {
            return CanCreateFolder(Constants.STRUCTLEVEL_04_HAUS);
        }

        // ------------------------------------------------------------------

        #endregion CreateFolder_Level4 (cmd_CreateFolder_Level4)

        // ==================================================================

        #region CreateFolder_Level5 (cmd_CreateFolder_Level5)

        public DelegateCommand cmd_CreateFolder_Level5 { get { return new DelegateCommand(CreateFolder_Level5, _CanCreateFolder_Level5); } }

        public void CreateFolder_Level5()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanCreateFolder_Level5)
                    CreateFolder(Constants.STRUCTLEVEL_05_RAUM);
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        public bool CanCreateFolder_Level5 { get { return _CanCreateFolder_Level5(); } }

        private bool _CanCreateFolder_Level5()
        {
            return CanCreateFolder(Constants.STRUCTLEVEL_05_RAUM);
        }

        // ------------------------------------------------------------------

        #endregion CreateFolder_Level5 (cmd_CreateFolder_Level5)

        // ==================================================================

        #region CreateFolder_Level6 (cmd_CreateFolder_Level6)

        public DelegateCommand cmd_CreateFolder_Level6 { get { return new DelegateCommand(CreateFolder_Level6, _CanCreateFolder_Level6); } }

        public void CreateFolder_Level6()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanCreateFolder_Level6)
                    CreateFolder(Constants.STRUCTLEVEL_06_SCHRANK);
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        public bool CanCreateFolder_Level6 { get { return _CanCreateFolder_Level6(); } }

        private bool _CanCreateFolder_Level6()
        {
            return CanCreateFolder(Constants.STRUCTLEVEL_06_SCHRANK);
        }

        // ------------------------------------------------------------------

        #endregion CreateFolder_Level6 (cmd_CreateFolder_Level6)

        // ==================================================================

        #region CreateFolder_Level7 (cmd_CreateFolder_Level7)

        public DelegateCommand cmd_CreateFolder_Level7 { get { return new DelegateCommand(CreateFolder_Level7, _CanCreateFolder_Level7); } }

        public void CreateFolder_Level7()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanCreateFolder_Level7)
                    CreateFolder(Constants.STRUCTLEVEL_07_AKTE);
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        public bool CanCreateFolder_Level7 { get { return _CanCreateFolder_Level7(); } }

        private bool _CanCreateFolder_Level7()
        {
            return CanCreateFolder(Constants.STRUCTLEVEL_07_AKTE);
        }

        // ------------------------------------------------------------------

        #endregion CreateFolder_Level7 (cmd_CreateFolder_Level7)

        // ==================================================================

        #region CreateFolder_Level8 (cmd_CreateFolder_Level8)

        public DelegateCommand cmd_CreateFolder_Level8 { get { return new DelegateCommand(CreateFolder_Level8, _CanCreateFolder_Level8); } }

        public void CreateFolder_Level8()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanCreateFolder_Level8)
                    CreateFolder(Constants.STRUCTLEVEL_08_VORGANG);
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        public bool CanCreateFolder_Level8 { get { return _CanCreateFolder_Level8(); } }

        private bool _CanCreateFolder_Level8()
        {
            return CanCreateFolder(Constants.STRUCTLEVEL_08_VORGANG);
        }

        // ------------------------------------------------------------------

        #endregion CreateFolder_Level8 (cmd_CreateFolder_Level8)

        // ==================================================================

        #region CreateFolder_Level9 (cmd_CreateFolder_Level9)

        public DelegateCommand cmd_CreateFolder_Level9 { get { return new DelegateCommand(CreateFolder_Level9, _CanCreateFolder_Level9); } }

        public void CreateFolder_Level9()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanCreateFolder_Level9)
                    CreateFolder(Constants.STRUCTLEVEL_09_DOKLOG);
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        public bool CanCreateFolder_Level9 { get { return _CanCreateFolder_Level9(); } }

        private bool _CanCreateFolder_Level9()
        {
            return CanCreateFolder(Constants.STRUCTLEVEL_09_DOKLOG);
        }

        // ------------------------------------------------------------------

        #endregion CreateFolder_Level9 (cmd_CreateFolder_Level9)

        // ==================================================================

        #region CreateFolder

        public void CreateFolder(int structlevel)
        {
            CreateFolder(structlevel, false, null, null, null);
        }

        public void CreateFolder(int structlevel, bool setaddressflag, IDocumentOrFolder parent)
        {
            CreateFolder(structlevel, setaddressflag, parent, null, null);
        }

        // TS 20.01.15 umbau auf callbackaction
        //public void CreateFolder(int structlevel, bool setaddressflag, IDocumentOrFolder parent, cmisTypeContainer typedef, Action callback)

        // TS 18.06.15 autosave logik generell umgebaut: es wird IMMER erst das Objekt ungespeichert zum Client geliefert
        // Autosave findet dann nur beim Wechsel des Objekts statt, nicht mehr sofort serverseitig
        public void CreateFolder(int structlevel, bool setaddressflag, IDocumentOrFolder parent, cmisTypeContainer typedef, CallbackAction callback)
        {
            bool autosave = DataAdapter.Instance.DataCache.Profile.Option_GetValueBoolean(CSEnumOptions.autosave_create);
            CreateFolder(structlevel, setaddressflag, parent, typedef, autosave, null);
        }

        // TS 18.06.15 autosave logik generell umgebaut (siehe oben)
        public void CreateFolder(int structlevel, bool setaddressflag, IDocumentOrFolder parent, cmisTypeContainer typedef, bool autosave, CallbackAction callback)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanCreateFolder(structlevel, parent, typedef, setaddressflag))
                {
                    // TS 18.06.15 autosave logik generell umgebaut (siehe oben)
                    autosave = false;

                    // TS 29.02.12
                    if (DataAdapter.Instance.DataCache.Objects(Workspace).Object_Selected.objectId.Length == 0)
                    {
                        DataAdapter.Instance.DataCache.Objects(Workspace).SetSelectedObject(DataAdapter.Instance.DataCache.Objects(Workspace).Root);
                        // die dummy objekte zur suche wegnehmen da ja nun ein echter ordner angelegt wird
                        DataAdapter.Instance.DataCache.Objects(Workspace).RemoveEmptyQueryObjects(true);
                    }

                    // parent holen
                    parent = CreateObjectGetParent(parent);
                    cmisObjectType cmisparent = parent.CMISObject;

                    // TS 11.04.13 parent mitgeben damit eine typunterscheidung stattfinden kann falls mehrere kindertypen erlaubt sind
                    if (typedef == null) typedef = DataAdapter.Instance.DataCache.Repository(Workspace).GetTypeContainerForStructLevel(structlevel, parent);

                    // TS 26.03.14 prüfung auf pflichtfelder, falls vorhanden dann kein autosave
                    if (autosave) autosave = CreateFolderCheckCanAutosave(typedef, setaddressflag);

                    // TS 23.04.15
                    // =============================================================================
                    // TS 26.06.17 defaultvalues jetzt IMMER ueber den server nicht nur bei autosave weil es sonst im aktenplan nicht funktioniert
                    //if (autosave) typedef = AddDefaultValuesToTypeContainer(typedef, setaddressflag);
                    typedef = AddDefaultValuesToTypeContainer(typedef, setaddressflag);

                    // TS 05.04.13
                    if (callback != null)
                    {
                        DataAdapter.Instance.DataProvider.CreateFolder(typedef, cmisparent, setaddressflag, autosave, DataAdapter.Instance.DataCache.Rights.UserPrincipal, new CallbackAction(CreateFolder_Done, callback));
                    }
                    else
                        DataAdapter.Instance.DataProvider.CreateFolder(typedef, cmisparent, setaddressflag, autosave, DataAdapter.Instance.DataCache.Rights.UserPrincipal, new CallbackAction(CreateFolder_Done));
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        private bool CreateFolderCheckCanAutosave(cmisTypeContainer typedef, bool setaddressflag)
        {
            bool canautosave = false;
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                List<string> requiredvalues = new List<string>();
                if (!setaddressflag)
                    requiredvalues = DataAdapter.Instance.DataCache.Profile.Profile_GetRequiredValues(Workspace, typedef.type.id);
                else
                    requiredvalues = DataAdapter.Instance.DataCache.Profile.Profile_GetRequiredValues(CSEnumProfileWorkspace.workspace_adressen, typedef.type.id);

                bool anyfound = false;
                if (requiredvalues != null && requiredvalues.Count > 0)
                {
                    // TS 27.04.15
                    Dictionary<string, string> defaultvalues = null;
                    if (setaddressflag)
                        defaultvalues = DataAdapter.Instance.DataCache.Profile.Profile_GetDefaultValues(CSEnumProfileWorkspace.workspace_adressen, typedef.type.id);
                    else
                        defaultvalues = DataAdapter.Instance.DataCache.Profile.Profile_GetDefaultValues(this.Workspace, typedef.type.id);

                    foreach (string value in requiredvalues)
                    {
                        // TS 27.04.15 IDX_DOKUMENTSYS_03 und IDX_DOKUMENTSYS_04 nicht mehr weglassen weil nun defaultvalues verfügbar sind
                        //if (!value.Equals(Constants.IDX_DOKUMENT_EINGDATUM.ToString()) && !value.Equals(Constants.IDX_DOKUMENT_DOKDATUM.ToString()))
                        //{
                        // TS 18.06.14
                        if ((setaddressflag && value.Contains(Constants.ADDRESSPROPERTY)) || (!setaddressflag && !value.Contains(Constants.ADDRESSPROPERTY)))
                        {
                            // TS 27.04.15 die defaultvalues dagegenprüfen
                            //anyfound = true;
                            //break;
                            bool found = false;
                            if (defaultvalues != null && defaultvalues.Count > 0)
                            {
                                foreach (string key in defaultvalues.Keys)
                                {
                                    if (key.Equals(value))
                                    {
                                        found = true;
                                        break;
                                    }
                                }
                                if (!found)
                                {
                                    anyfound = true;
                                    break;
                                }
                            }
                            else
                            {
                                anyfound = true;
                                break;
                            }
                        }
                        //}
                    }
                    //if (anyfound)
                    //    canautosave = false;
                }
                canautosave = true;
                if (anyfound) canautosave = false;
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
            return canautosave;
        }

        // ------------------------------------------------------------------
        public void CreateFolder_Done() { CreateFolder_Done(null); }

        public void CreateFolder_Done(CallbackAction callback)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (DataAdapter.Instance.DataCache.ResponseStatus.success)
                {
                    IDocumentOrFolder createdfolder = DataAdapter.Instance.DataCache.Objects(Workspace).ObjectList[DataAdapter.Instance.DataCache.Objects(Workspace).ObjectList.Count - 1];
                    DataAdapter.Instance.DataCache.Info.ListView_ForcePageSelection = true;
                    this.SetSelectedObject(createdfolder.objectId);

                    // TS 17.12.13 wen automatisches speichern gewünscht war aber nicht erfolgt ist (weil anschliessend keine rechte): meldung ausgeben
                    // TS 10.03.14 umgestellt auf andere logik: autosave erst bei wechsel des parents
                    // TS 26.03.14 wieder reingenommen

                    // TS 18.06.15 autosave logik generell umgebaut (siehe oben)
                    bool autosave = false;
                    if (autosave && createdfolder.isNotCreated)
                        DisplayWarnMessage(LocalizationMapper.Instance["msg_autosave_folder_notdone"]);

                    // Check auto-numbering
                    //if (!DataAdapter.Instance.DataCache.Profile.Option_GetValueBoolean(CSEnumOptions.autosave_create))
                    //{
                    //    CheckAutoNumberForObject(createdfolder);
                    //}

                    // Default-Sortorder, if wanted
                    if (DataAdapter.Instance.DataCache.Profile.Option_GetValueBoolean(CSEnumOptions.treesortorderfix))
                    {
                        createdfolder.Folge = "-2";
                    }

                    // Set hasChildFolders Property
                    if(createdfolder.ParentFolder != null && !createdfolder.ParentFolder.hasChildFolders)
                    {
                        createdfolder.ParentFolder.SetPropertyValue(CSEnumCmisProperties.hasChildFolders.ToString(), "true", true);
                        createdfolder.ParentFolder.SetPropertyValue(CSEnumCmisProperties.childrenHasMore.ToString(), "false", true);
                    }

                }

                // TS 25.11.13
                if (callback != null)
                    callback.Invoke();
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        // TS 30.01.15
        //public bool CanCreateFolder(int structlevel) { return CanCreateFolder(structlevel, null, null); }
        //public bool CanCreateFolder(int structlevel, IDocumentOrFolder parent, cmisTypeContainer typedef) { return CanCreateFolder(structlevel, parent, typedef); }
        private bool CanCreateFolder(int structlevel) { return CanCreateFolder(structlevel, null, null, false); }

        private bool CanCreateFolder(int structlevel, IDocumentOrFolder parent, cmisTypeContainer typedef)
        {
            return CanCreateFolder(structlevel, parent, typedef, false);
        }

        public bool CanCreateFolder(int structlevel, bool isaddress)
        {
            return CanCreateFolder(structlevel, null, null, isaddress);
        }

        public bool CanCreateFolder(int structlevel, IDocumentOrFolder parent, cmisTypeContainer typedef, bool isaddress)
        {
            bool cancreate = false;

            if (DataAdapter.Instance.DataCache.ApplicationFullyInit)
            {
                parent = CreateObjectGetParent(parent);

                if (typedef == null)
                    cancreate = parent.canCreateObjectLevel(structlevel);
                else
                    cancreate = parent.canCreateObjectTypeId(typedef.type.id);

                // TS 30.01.15 adressflags prüfen
                if (cancreate && structlevel > Constants.STRUCTLEVEL_07_AKTE)
                {
                    if (isaddress)
                        cancreate = parent.isAddressObject;
                    else
                    {
                        // TS 05.02.15
                        //cancreate = !parent.isAddressObject;
                        cancreate = parent.isPureOfficeObject;
                    }
                }
            }
            return cancreate;
        }

        ///// <summary>
        ///// Checks all siblings of the new object for the same autonumber-id if available
        ///// Increments the autonumber-id
        ///// </summary>
        ///// <param name="newObj"></param>
        //private void CheckAutoNumberForObject(IDocumentOrFolder newObj)
        //{
        //    string autoPropName = newObj[CSEnumCmisProperties.tmp_autonumbproperty];
        //    if(autoPropName != null && autoPropName.Length > 0)
        //    {
        //        string objAutoNumb = newObj[autoPropName];
        //        bool siblingFound = true;
        //        if (objAutoNumb != null && objAutoNumb.Length > 0)
        //        {
        //            // Loop until no sibling with the same autonumber-value has been found
        //            while (siblingFound)
        //            {
        //                siblingFound = false;
        //                objAutoNumb = newObj[autoPropName];
        //                foreach (IDocumentOrFolder objSibling in DataAdapter.Instance.DataCache.Objects(Workspace).GetDescendants(newObj.ParentFolder))
        //                {
        //                    if (objSibling != newObj)
        //                    {
        //                        string siblingAutoNumb = objSibling[autoPropName];
        //                        if (siblingAutoNumb != null && siblingAutoNumb.Length > 0)
        //                        {
        //                            // Increment the autonumber-value
        //                            if (siblingAutoNumb.Equals(objAutoNumb))
        //                            {
        //                                // TS 17.08.16
        //                                //newObj[autoPropName] = (int.Parse(objAutoNumb) + 1).ToString();

        //                                string tmpautonumb = objAutoNumb;
        //                                int pos = tmpautonumb.Length - 1;
        //                                int length = 0;
        //                                while (IsNumeric(tmpautonumb.Substring(pos, 1)))
        //                                {
        //                                    pos--;
        //                                    length++;
        //                                }
        //                                string realnumber = tmpautonumb.Substring(pos + 1, length);
        //                                string picture = tmpautonumb.Substring(0, pos + 1);
        //                                int newnumber = (int.Parse(realnumber)) + 1;
        //                                newObj[autoPropName] = picture + newnumber.ToString();

        //                                siblingFound = true;
        //                                break;
        //                            }
        //                        }
        //                    }
        //                }
        //            }
        //        }
        //    }
        //}

        private bool IsNumeric(string s)
        {
            float output;
            return float.TryParse(s, out output);
        }

        private IDocumentOrFolder CreateObjectGetParent(IDocumentOrFolder parent)
        {
            if (parent == null)
                // TS 26.11.13
                //parent = DataAdapter.Instance.DataCache.Objects(Workspace).Object_Selected;
                parent = DataAdapter.Instance.DataCache.Objects(Workspace).Folder_Selected;
            if (parent.objectId.Length == 0)
                parent = DataAdapter.Instance.DataCache.Objects(Workspace).Root;
            return parent;
        }

        // ------------------------------------------------------------------

        #endregion CreateFolder

        // ==================================================================

        #region CreateOrUpdateDocAnno

        public void CreateOrUpdateDocAnno(CSAnnotationFile annofile)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanCreateDocAnno)
                {
                    // annofile suchen
                    IDocumentOrFolder annodoc = _FindAnnofile();
                    if (annodoc == null)
                    {
                        IDocumentOrFolder dummy = DataAdapter.Instance.DataCache.Objects(Workspace).Folder_Selected;
                        cmisObjectType parent = dummy.CMISObject;
                        IDocumentOrFolder currentdoc = DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected;
                        cmisTypeContainer doctypedefinition = DataAdapter.Instance.DataCache.Repository(Workspace).GetTypeContainerForStructLevel(10);

                        // TS 03.12.13
                        bool autosave = DataAdapter.Instance.DataCache.Profile.Option_GetValueBoolean(CSEnumOptions.autosave_create);
                        // TS 25.02.14 callback dazu wenn autosave weil dann vom server ein neues doklog kommt und die collection füpr die thumbs verschwindet !
                        //DataAdapter.Instance.DataProvider.CreateAnnotationDocument(doctypedefinition, parent, annofile, currentdoc.objectId, autosave, DataAdapter.Instance.DataCache.Rights.UserPrincipal);
                        CallbackAction callback = null;
                        if (autosave) callback = new CallbackAction(CreateDocAnno_Done);
                        DataAdapter.Instance.DataProvider.CreateAnnotationDocument(doctypedefinition, parent, annofile, currentdoc.objectId, autosave, DataAdapter.Instance.DataCache.Rights.UserPrincipal, callback);
                    }
                    else
                    {
                        // TS 25.02.14 reihenfolge muss getauscht werden da die property: annodoc.AnnotationFile durch autosave bereits ein save feuert während noch der update aktiv ist
                        // besser mit callback
                        //annodoc.AnnotationFile = annofile;

                        // TS 25.02.14 callback
                        //DataAdapter.Instance.DataProvider.UpdateAnnotationFile(annodoc.CMISObject, annofile, DataAdapter.Instance.DataCache.Rights.UserPrincipal);
                        CallbackAction callback = new CallbackAction(UpdateDocAnno_Done, annodoc, annofile);
                        DataAdapter.Instance.DataProvider.UpdateAnnotationFile(annodoc.CMISObject, annofile, DataAdapter.Instance.DataCache.Rights.UserPrincipal, callback);
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        public bool CanCreateDocAnno { get { return _CanCreateDocAnno(); } }

        private bool _CanCreateDocAnno()
        {
            bool cancreate = CanCreateDocuments && DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected.objectId.Length > 0;
            return cancreate;
        }

        private void CreateDocAnno_Done()
        {
            if (DataAdapter.Instance.DataCache.ResponseStatus.success)
            {
                IDocumentOrFolder createddoc = DataAdapter.Instance.DataCache.Objects(Workspace).ObjectList[DataAdapter.Instance.DataCache.Objects(Workspace).ObjectList.Count - 1];
                if (createddoc.isDocument)
                {
                    RefreshDZCollectionsAfterProcess(createddoc, null);
                }
                //this.SetSelectedObject(createddoc.objectId);
                // TS 17.12.13 wen automatisches speichern gewünscht war aber nicht erfolgt ist (weil anschliessend keine rechte): meldung ausgeben
                // TS 10.03.14 umgestellt auf andere logik: autosave erst bei wechsel des parents
                // TS 26.03.14 wieder reingenommen
                bool autosave = DataAdapter.Instance.DataCache.Profile.Option_GetValueBoolean(CSEnumOptions.autosave_create);
                if (autosave && createddoc.isNotCreated)
                    DisplayWarnMessage(LocalizationMapper.Instance["msg_autosave_document_notdone"]);
            }
        }

        private void UpdateDocAnno_Done(object annodoc, object annofile)
        {
            if (DataAdapter.Instance.DataCache.ResponseStatus.success)
            {
                ((IDocumentOrFolder)annodoc).AnnotationFile = (CSAnnotationFile)annofile;
            }
        }

        private IDocumentOrFolder _FindAnnofile()
        {
            IDocumentOrFolder ret = null;
            foreach (IDocumentOrFolder obj in DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected.ChildObjects)
            {
                if (obj.CopyType == CSEnumDocumentCopyTypes.COPYTYPE_OVERLAY)
                {
                    ret = obj;
                    break;
                }
            }
            return ret;
        }

        #endregion CreateOrUpdateDocAnno

        // ==================================================================

        #region CreateWiedervorlage (cmd_CreateWiedervorlage) => gibts nicht mehr

        //public DelegateCommand cmd_CreateWiedervorlage { get { return new DelegateCommand(CreateWiedervorlage, _CanCreateWiedervorlage); } }
        //public void CreateWiedervorlage()
        //{
        //    if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
        //    try
        //    {
        //        if (CanCreateWiedervorlage)
        //        {
        //            // TS 24.02.14 umbau auf "Aufgabe"
        //            //_CreateFoldersInternal(CSEnumInternalObjectType.WDVL, CSEnumProfileWorkspace.workspace_termin, Constants.STRUCTLEVEL_09_DOKLOG);
        //            _CreateFoldersInternal(CSEnumInternalObjectType.WDVL, CSEnumProfileWorkspace.workspace_aufgabe, Constants.STRUCTLEVEL_09_DOKLOG);
        //        }
        //    }
        //    catch (Exception e) { Log.Log.Error(e); }
        //    if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        //}
        //// ------------------------------------------------------------------
        //public bool CanCreateWiedervorlage { get { return _CanCreateWiedervorlage(); } }
        //// TS 24.02.14 umbau auf aufgabe
        ////private bool _CanCreateWiedervorlage() { return _CanCreateFoldersInternal(CSEnumInternalObjectType.WDVL, CSEnumProfileWorkspace.workspace_termin, Constants.STRUCTLEVEL_09_DOKLOG); }
        //private bool _CanCreateWiedervorlage() { return _CanCreateFoldersInternal(CSEnumInternalObjectType.WDVL, CSEnumProfileWorkspace.workspace_aufgabe, Constants.STRUCTLEVEL_09_DOKLOG); }
        // ------------------------------------------------------------------

        #endregion CreateWiedervorlage (cmd_CreateWiedervorlage) => gibts nicht mehr

        // ==================================================================

        #region CreateTermin (cmd_CreateTermin)

        public DelegateCommand cmd_CreateTermin { get { return new DelegateCommand(CreateTermin, _CanCreateTermin); } }

        public void CreateTermin()
        {
            CreateTermin(DataAdapter.Instance.DataCache.Objects(Workspace).Objects_Selected);
        }

        public void CreateTermin(List<IDocumentOrFolder> selectedobjects)
        {
            CreateTermin(selectedobjects, true);
        }

        public void CreateTermin(List<IDocumentOrFolder> selectedobjects, bool showdialog)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (_CanCreateTermin(selectedobjects))
                    _CreateFoldersInternal(CSEnumInternalObjectType.Termin, CSEnumProfileWorkspace.workspace_termin, Constants.STRUCTLEVEL_09_DOKLOG, selectedobjects, showdialog);
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        public bool CanCreateTermin { get { return _CanCreateTermin(); } }

        private bool _CanCreateTermin()
        {
            return _CanCreateTermin(DataAdapter.Instance.DataCache.Objects(Workspace).Objects_Selected);
        }

        public bool _CanCreateTermin(List<IDocumentOrFolder> selectedobjects)
        {
            return _CanCreateFoldersInternal(CSEnumInternalObjectType.Termin, CSEnumProfileWorkspace.workspace_termin, Constants.STRUCTLEVEL_09_DOKLOG, selectedobjects);
        }

        // ------------------------------------------------------------------

        #endregion CreateTermin (cmd_CreateTermin)

        // ==================================================================

        #region CreateAufgabe (cmd_CreateAufgabe)

        public DelegateCommand cmd_CreateAufgabe { get { return new DelegateCommand(CreateAufgabe, _CanCreateAufgabe); } }

        public void CreateAufgabe()
        {
            CreateAufgabe(DataAdapter.Instance.DataCache.Objects(Workspace).Objects_Selected);
        }

        public void CreateAufgabe(List<IDocumentOrFolder> selectedobjects)
        {
            CreateAufgabe(selectedobjects, true);
        }

        public void CreateAufgabe(List<IDocumentOrFolder> selectedobjects, bool showdialog)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (_CanCreateAufgabe(selectedobjects) && selectedobjects.Count > 0)
                {
                    // Send a Push-Notification to the EDesktop to create a task
                    CSRCPushPullUIRouting routing = new CSRCPushPullUIRouting();
                    CSEnumRCPushUIListeners?[] receiver = new CSEnumRCPushUIListeners?[1];
                    receiver[0] = CSEnumRCPushUIListeners.HTML_CLIENT_DESKTOP;
                    routing.sendto_componentorvirtualids = receiver;

                    List<string> objectidlist = new List<string>();
                    foreach (IDocumentOrFolder doc in selectedobjects)
                    {
                        objectidlist.Add("2C_" + doc.RepositoryId + "_" + doc.objectId);
                    }

                    // If the HTML-Desktop is already open, send a push-command
                    if (Init.ViewManager.IsHTMLDesktopOpen())
                    {
                        DataAdapter.Instance.DataProvider.RCPush_UI(routing, CSEnumRCPushCommands.createaufgabe, objectidlist, null, DataAdapter.Instance.DataCache.Rights.UserPrincipal);
                    }else
                    {
                        // Open the desktop-tool                        
                        Toggle_ShowDesk(CSEnumRCPushCommands.createaufgabe.ToString() + "=" + objectidlist[0]);
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        public bool CanCreateAufgabe { get { return _CanCreateAufgabe(); } }

        private bool _CanCreateAufgabe()
        {
            return _CanCreateAufgabe(DataAdapter.Instance.DataCache.Objects(Workspace).Objects_Selected);
        }

        public bool _CanCreateAufgabe(List<IDocumentOrFolder> selectedobjects)
        {
            return _CanCreateFoldersInternal(CSEnumInternalObjectType.Aufgabe, CSEnumProfileWorkspace.workspace_aufgabe, Constants.STRUCTLEVEL_09_DOKLOG, selectedobjects) && ViewManager.IsHTMLDesktop;
        }

        // ------------------------------------------------------------------

        #endregion CreateAufgabe (cmd_CreateAufgabe)

        // ==================================================================

        #region CreateAufgabePublic (cmd_CreateAufgabePublic)

        public DelegateCommand cmd_CreateAufgabePublic { get { return new DelegateCommand(CreateAufgabePublic, _CanCreateAufgabePublic); } }

        public void CreateAufgabePublic()
        {
            CreateAufgabePublic(DataAdapter.Instance.DataCache.Objects(Workspace).Objects_Selected);
        }

        public void CreateAufgabePublic(List<IDocumentOrFolder> selectedobjects)
        {
            CreateAufgabePublic(selectedobjects, true);
        }

        public void CreateAufgabePublic(List<IDocumentOrFolder> selectedobjects, bool showdialog)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (_CanCreateAufgabePublic(selectedobjects) && selectedobjects.Count > 0)
                {
                    // Send a Push-Notification to the EDesktop to create a task
                    CSRCPushPullUIRouting routing = new CSRCPushPullUIRouting();
                    CSEnumRCPushUIListeners?[] receiver = new CSEnumRCPushUIListeners?[1];
                    receiver[0] = CSEnumRCPushUIListeners.HTML_CLIENT_DESKTOP;
                    routing.sendto_componentorvirtualids = receiver;

                    List<string> objectidlist = new List<string>();
                    foreach (IDocumentOrFolder doc in selectedobjects)
                    {
                        objectidlist.Add("2C_" + doc.RepositoryId + "_" + doc.objectId);
                    }

                    // If the HTML-Desktop is already open, send a push-command
                    if (Init.ViewManager.IsHTMLDesktopOpen())
                    {                        
                        DataAdapter.Instance.DataProvider.RCPush_UI(routing, CSEnumRCPushCommands.createaufgabepublic, objectidlist, null, DataAdapter.Instance.DataCache.Rights.UserPrincipal);
                    }
                    else
                    {
                        // Open the desktop-tool                        
                        Toggle_ShowDesk(CSEnumRCPushCommands.createaufgabepublic.ToString() + "=" + objectidlist[0]);
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        public bool CanCreateAufgabePublic { get { return _CanCreateAufgabePublic(); } }

        private bool _CanCreateAufgabePublic()
        {
            return _CanCreateAufgabePublic(DataAdapter.Instance.DataCache.Objects(Workspace).Objects_Selected);
        }

        public bool _CanCreateAufgabePublic(List<IDocumentOrFolder> selectedobjects)
        {
            return _CanCreateFoldersInternal(CSEnumInternalObjectType.AufgabePublic, CSEnumProfileWorkspace.workspace_aufgabe, Constants.STRUCTLEVEL_09_DOKLOG, selectedobjects) && ViewManager.IsHTMLDesktop;
        }

        // ------------------------------------------------------------------

        #endregion CreateAufgabePublic (cmd_CreateAufgabePublic)

        // ==================================================================

        #region CreateStapel (cmd_CreateStapel)

        public DelegateCommand cmd_CreateStapel { get { return new DelegateCommand(CreateStapel, _CanCreateStapel); } }

        public void CreateStapel()
        {
            CreateStapel(DataAdapter.Instance.DataCache.Objects(Workspace).Objects_Selected);
        }

        public void CreateStapel(List<IDocumentOrFolder> selectedobjects)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (_CanCreateStapel(selectedobjects))
                {
                    // TS 19.02.14 umbau wg. datenmodellwechsel internes archiv
                    //_CreateFoldersInternal(CSEnumInternalObjectType.Stapel, CSEnumProfileWorkspace.workspace_stapel, Constants.STRUCTLEVEL_09_DOKLOG);
                    IDocumentOrFolder targetobject = null;
                    // TS 28.02.14
                    //string targetobjectid = DataAdapter.Instance.DataCache.Profile.Profile_GetInternalRootId(Constants.INTERNAL_DOCTYPE_STAPEL);
                    string targetobjectid = DataAdapter.Instance.DataCache.Profile.Profile_GetInternalRootId(CSEnumInternalObjectType.Stapel.ToString());
                    if (targetobjectid != null && targetobjectid.Length > 0)
                        targetobject = DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_stapel).GetObjectById(targetobjectid);
                    Move(selectedobjects, targetobject);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        public bool CanCreateStapel { get { return _CanCreateStapel(); } }

        // TS 19.02.14 umbau wg. datenmodellwechsel internes archiv
        //private bool _CanCreateStapel() { return _CanCreateFoldersInternal(CSEnumInternalObjectType.Stapel, CSEnumProfileWorkspace.workspace_stapel, Constants.STRUCTLEVEL_09_DOKLOG); }
        private bool _CanCreateStapel() { return _CanCreateStapel(DataAdapter.Instance.DataCache.Objects(Workspace).Objects_Selected); }

        public bool _CanCreateStapel(List<IDocumentOrFolder> selectedobjects)
        {
            bool ret = false;
            IDocumentOrFolder targetobject = null;
            // TS 28.02.14
            //string targetobjectid = DataAdapter.Instance.DataCache.Profile.Profile_GetInternalRootId(Constants.INTERNAL_DOCTYPE_STAPEL);
            string targetobjectid = DataAdapter.Instance.DataCache.Profile.Profile_GetInternalRootId(CSEnumInternalObjectType.Stapel.ToString());
            if (targetobjectid != null && targetobjectid.Length > 0)
                targetobject = DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_stapel).GetObjectById(targetobjectid);
            ret = targetobject != null && selectedobjects != null && selectedobjects.Count > 0;
            // TS 24.02.14 zunächst nur dokumente erlauben
            if (ret)
            {
                foreach (IDocumentOrFolder child in selectedobjects)
                {
                    if (child.structLevel != 9)
                    {
                        ret = false;
                        break;
                    }
                }
            }
            if (ret) ret = CanMove(selectedobjects, targetobject);
            return ret;
        }

        // ------------------------------------------------------------------

        #endregion CreateStapel (cmd_CreateStapel)

        // ==================================================================

        #region CreateAblage (cmd_CreateAblage)

        public DelegateCommand cmd_CreateAblage { get { return new DelegateCommand(CreateAblage, _CanCreateAblage); } }

        public void CreateAblage()
        {
            CreateAblage(DataAdapter.Instance.DataCache.Objects(Workspace).Objects_Selected);
        }

        public void CreateAblage(List<IDocumentOrFolder> selectedobjects)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (_CanCreateAblage(selectedobjects))
                {
                    // TS 19.02.14 umbau wg. datenmodellwechsel internes archiv
                    //_CreateFoldersInternal(CSEnumInternalObjectType.Stapel, CSEnumProfileWorkspace.workspace_stapel, Constants.STRUCTLEVEL_09_DOKLOG);
                    IDocumentOrFolder targetobject = null;
                    // TS 28.02.14
                    //string targetobjectid = DataAdapter.Instance.DataCache.Profile.Profile_GetInternalRootId(Constants.INTERNAL_DOCTYPE_STAPEL);
                    string targetobjectid = DataAdapter.Instance.DataCache.Profile.Profile_GetInternalRootId(CSEnumInternalObjectType.Ablage.ToString());
                    if (targetobjectid != null && targetobjectid.Length > 0)
                        targetobject = DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_ablage).GetObjectById(targetobjectid);
                    Move(selectedobjects, targetobject);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        public bool CanCreateAblage { get { return _CanCreateAblage(); } }

        private bool _CanCreateAblage()
        {
            return _CanCreateAblage(DataAdapter.Instance.DataCache.Objects(Workspace).Objects_Selected);
        }

        public bool _CanCreateAblage(List<IDocumentOrFolder> selectedobjects)
        {
            bool ret = false;
            IDocumentOrFolder targetobject = null;
            string targetobjectid = DataAdapter.Instance.DataCache.Profile.Profile_GetInternalRootId(CSEnumInternalObjectType.Ablage.ToString());
            if (targetobjectid != null && targetobjectid.Length > 0)
                targetobject = DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_ablage).GetObjectById(targetobjectid);
            ret = targetobject != null && selectedobjects != null && selectedobjects.Count > 0;
            // TS 24.02.14 zunächst nur dokumente erlauben
            if (ret)
            {
                foreach (IDocumentOrFolder child in selectedobjects)
                {
                    if (child.structLevel != 9)
                    {
                        ret = false;
                        break;
                    }
                }
            }
            if (ret) ret = CanMove(selectedobjects, targetobject);
            return ret;
        }

        // ------------------------------------------------------------------

        #endregion CreateAblage (cmd_CreateAblage)

        // ==================================================================

        #region CreateRichText (cmd_CreateRichText)

        public DelegateCommand cmd_CreateRichText { get { return new DelegateCommand(CreateRichText, CanCreateRichText); } }

        public void CreateRichText()
        {
            CreateRichText(false);
        }

        public void CreateRichText(bool setaddressflag)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanCreateRichText(setaddressflag))
                {
                    // TS 16.03.15
                    //IDocumentOrFolder currentfolder = Page.PageBindingAdapter.DataCacheObjects.Folder_Selected;
                    IDocumentOrFolder currentfolder = DataAdapter.Instance.DataCache.Objects(this.Workspace).Folder_Selected;

                    // TS 08.05.14 autosave kurzzeitig abschalten
                    // TS 20.01.15 autosave = false auch als parameter mitgeben
                    bool wasautosave = DataAdapter.Instance.DataCache.Profile.Option_GetValueBoolean(CSEnumOptions.autosave_create);
                    DataAdapter.Instance.DataCache.Profile.Option_SetValueBoolean(CSEnumOptions.autosave_create, false);

                    // TS 16.03.15 wenn bereits auf doklog positioniert dann weiter
                    if (currentfolder.structLevel == Constants.STRUCTLEVEL_09_DOKLOG)
                    {
                        CreateRichTextCreateFolderDone(wasautosave);
                    }
                    else
                    {
                        // TS 18.06.15 autosave generell umgebaut
                        CallbackAction callback = new CallbackAction(CreateRichTextCreateFolderDone, wasautosave);
                        CreateFolder(Constants.STRUCTLEVEL_09_DOKLOG, setaddressflag, currentfolder, null, false, callback);
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        private void CreateRichTextCreateFolderDone(bool wasautosave)
        {
            // ein paar indizes setzen
            if (DataAdapter.Instance.DataCache.ResponseStatus.success)
            {
                IDocumentOrFolder currentfolder = DataAdapter.Instance.DataCache.Objects(this.Workspace).Folder_Selected;
                // TS 16.03.15
                //if (currentfolder.structLevel == Constants.STRUCTLEVEL_09_DOKLOG)
                //{
                //    currentfolder[Statics.Constants.IDX_ADR_DOKUMENT_KONTAKTDATUM] = DateFormatHelper.GetBasicDate();
                //    currentfolder[Statics.Constants.IDX_ADR_DOKUMENT_BEMERKUNG] = Localization.localstring.D_ADR_KONTAKT_NOTIZ;
                //}
                // neues dokument anlegen
                cmisTypeContainer doctypedefinition = DataAdapter.Instance.DataCache.Repository(this.Workspace).GetTypeContainerForStructLevel(10);
                IDocumentOrFolder dummy = new DocumentOrFolder(this.Workspace);
                // current folder mitgeben
                // TS 20.01.15 wasautosave mitgeben new CallbackAction(EditRichText_ShowDialog, currentfld));
                DataAdapter.Instance.DataProvider.CreateRichTextDocument(doctypedefinition, currentfolder.CMISObject, DataAdapter.Instance.DataCache.Rights.UserPrincipal,
                    new CallbackAction(EditRichText_ShowDialog, currentfolder, wasautosave));
            }
            else if (wasautosave)
                DataAdapter.Instance.DataCache.Profile.Option_SetValueBoolean(CSEnumOptions.autosave_create, true);
        }

        private bool CanCreateRichText()
        {
            return CanCreateRichText(false);
        }

        public bool CanCreateRichText(bool isaddress)
        {
            bool ret = DataAdapter.Instance.DataCache.ApplicationFullyInit;
            if (ret)
            {
                ret = false;
                IDocumentOrFolder currentfolder = DataAdapter.Instance.DataCache.Objects(this.Workspace).Folder_Selected;
                if (currentfolder.structLevel == Constants.STRUCTLEVEL_07_AKTE && currentfolder.canCreateObjectLevel(Statics.Constants.STRUCTLEVEL_09_DOKLOG))
                {
                    // TS 03.07.15 dieser mist mit den adressen zieht sich immer weiter, das muss demnächst generell komplett anders gelöst werden
                    //ret = true;
                    if (isaddress)
                        ret = currentfolder.isAddressObject;
                    else
                    {
                        ret = currentfolder.isPureOfficeObject;
                    }
                }
                if (currentfolder.structLevel == Constants.STRUCTLEVEL_08_VORGANG && currentfolder.canCreateObjectLevel(Statics.Constants.STRUCTLEVEL_09_DOKLOG))
                {
                    // TS 03.07.15 dieser mist mit den adressen zieht sich immer weiter, das muss demnächst generell komplett anders gelöst werden
                    //ret = true;
                    if (isaddress)
                        ret = currentfolder.isAddressObject;
                    else
                    {
                        ret = currentfolder.isPureOfficeObject;
                    }                   
                }
                if (currentfolder.structLevel == Constants.STRUCTLEVEL_09_DOKLOG)
                {
                    // TS 03.07.15 dieser mist mit den adressen zieht sich immer weiter, das muss demnächst generell komplett anders gelöst werden
                    if (isaddress)
                        ret = currentfolder.isAddressObject;
                    else
                    {
                        ret = currentfolder.isPureOfficeObject;
                    }

                    // TS 03.07.15 dieser mist mit den adressen zieht sich immer weiter, das muss demnächst generell komplett anders gelöst werden
                    if (ret)
                    {
                        bool anyrichtextfound = false;
                        foreach (IDocumentOrFolder child in currentfolder.ChildDocuments)
                        {
                            if (child.isRichTextDoc)
                            {
                                anyrichtextfound = true;
                                break;
                            }
                        }
                        ret = !anyrichtextfound;
                    }

                    if (currentfolder.ACL != null)
                    {
                        bool readOnly = currentfolder.hasCreatePermission();
                        ret = ret && readOnly;
                    }
                }
            }
            return ret;
        }

        //private bool CanCreateFolder(int structlevel) { return CanCreateFolder(structlevel, null, null, false); }
        //private bool CanCreateFolder(int structlevel, IDocumentOrFolder parent, cmisTypeContainer typedef) { return CanCreateFolder(structlevel, parent, typedef, false); }
        //public bool CanCreateFolder(int structlevel, bool isaddress) { return CanCreateFolder(structlevel, null, null, isaddress); }
        //public bool CanCreateFolder(int structlevel, IDocumentOrFolder parent, cmisTypeContainer typedef, bool isaddress)

        #endregion CreateRichText (cmd_CreateRichText)

        // ==================================================================

        #region EditRichText (cmd_EditRichText)

        public DelegateCommand cmd_EditRichText { get { return new DelegateCommand(EditRichText, _CanEditRichText); } }

        public void EditRichText()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanEditRichText)
                {
                    IDocumentOrFolder currentfolder = DataAdapter.Instance.DataCache.Objects(this.Workspace).Folder_Selected;
                    IDocumentOrFolder richtextdoc = null;

                    // TS 23.06.17
                    // if (DataAdapter.Instance.DataCache.Objects(this.Workspace).Document_Selected.isRichTextDoc)
                    //      richtextdoc = DataAdapter.Instance.DataCache.Objects(this.Workspace).Document_Selected;
                    if (DataAdapter.Instance.DataCache.Objects(this.Workspace).Object_Selected.isRichTextDoc)
                        richtextdoc = DataAdapter.Instance.DataCache.Objects(this.Workspace).Object_Selected;
                    else
                        richtextdoc = EditRichText_GetNewestDoc(currentfolder);

                    if (richtextdoc.isNotCreated || richtextdoc.isEdited || richtextdoc.canCheckIn)
                    {
                        // wenn noch nicht gespeichert oder bereits neue version und nicht eingecheckt dann dieses dok verwenden
                        EditRichText_ShowDialog(richtextdoc);
                    }
                    else
                    {
                        // sonst eine neue version erzeugen
                        CreateDocIVersionFrom(richtextdoc, new CallbackAction(EditRichText_ShowDialog, currentfolder));
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public bool CanEditRichText { get { return _CanEditRichText(); } }

        private bool _CanEditRichText()
        {
            IDocumentOrFolder richtextdoc = null;

            // TS 23.06.17
            //if (DataAdapter.Instance.DataCache.Objects(this.Workspace).Document_Selected.isRichTextDoc)
            //    richtextdoc = DataAdapter.Instance.DataCache.Objects(this.Workspace).Document_Selected;
            if (DataAdapter.Instance.DataCache.Objects(this.Workspace).Object_Selected.isRichTextDoc)
                richtextdoc = DataAdapter.Instance.DataCache.Objects(this.Workspace).Object_Selected;
            else
                richtextdoc = EditRichText_GetNewestDoc(DataAdapter.Instance.DataCache.Objects(this.Workspace).Folder_Selected);

            // HIEWR NOCH ZUSAETZLICH PRUEFEN:
            //return DataAdapter.Instance.DataCache.ApplicationFullyInit
            //    && DataAdapter.Instance.DataCache.Objects(this.Workspace).Folder_Selected.structLevel == Constants.STRUCTLEVEL_09_DOKLOG
            //    && richtextdoc != null;
            // ja => wenn (richtextdoc.isNotCreated || richtextdoc.isEdited || richtextdoc.canCheckIn)
            // sonst CanCreateDocIVersionFrom prüfen

            bool ret = DataAdapter.Instance.DataCache.ApplicationFullyInit
                && DataAdapter.Instance.DataCache.Objects(this.Workspace).Folder_Selected.structLevel == Constants.STRUCTLEVEL_09_DOKLOG
                && richtextdoc != null;
            if (ret)
            {
                // entweder es ist noch nicht gespeichert oder so, dann kann es direkt geaendert werden
                ret = (richtextdoc.isNotCreated || richtextdoc.isEdited || richtextdoc.canCheckIn);
                // oder es muss eine neue version erzeugt werden
                if (!ret)
                {
                    ret = _CanCreateDocIVersionFrom(richtextdoc);
                }
            }
            return ret;
        }

        private IDocumentOrFolder EditRichText_GetNewestDoc(IDocumentOrFolder parentfolder)
        {
            IDocumentOrFolder ret = null;
            List<IDocumentOrFolder> newest = parentfolder.GetNewestChildDokVersions(false);
            if (newest != null && newest.Count > 0)
            {
                foreach (IDocumentOrFolder rtchild in newest)
                {
                    if (rtchild.isRichTextDoc)
                    {
                        ret = rtchild;
                        break;
                    }
                }
            }
            return ret;
        }

        private void EditRichText_ShowDialog(object richtextdocorfolder)
        {
            EditRichText_ShowDialog(richtextdocorfolder, false);
        }

        private void EditRichText_ShowDialog(object richtextdocorfolder, object wasautosave)
        {
            if (richtextdocorfolder != null)
            {
                IDocumentOrFolder richtextdoc = null;
                IDocumentOrFolder docorfolder = ((IDocumentOrFolder)richtextdocorfolder);
                if (docorfolder.isFolder)
                {
                    // wenn ein ordner übergeben wurde dann wurde das dokument gerade erst erstellt oder ausgecheckt, daher prüfen
                    if (DataAdapter.Instance.DataCache.ResponseStatus.success)
                        richtextdoc = EditRichText_GetNewestDoc(docorfolder);
                }
                else if (docorfolder.isDocument)
                    richtextdoc = docorfolder;

                if (richtextdoc != null && richtextdoc.objectId.Length > 0 && richtextdoc.isRichTextDoc)
                {
                    SL2C_Client.View.Dialogs.dlgEditRichText child = new SL2C_Client.View.Dialogs.dlgEditRichText((IDocumentOrFolder)richtextdoc);
                    DialogHandler.Show_Dialog(child);
                }
            }
            // TS 20.01.15
            if (wasautosave != null && ((bool)wasautosave) == true) DataAdapter.Instance.DataCache.Profile.Option_SetValueBoolean(CSEnumOptions.autosave_create, true);
        }

        #endregion EditRichText (cmd_EditRichText)

        // ==================================================================

        #region Cut (cmd_Cut)

        public DelegateCommand cmd_Cut { get { return new DelegateCommand(Cut, _CanCut); } }

        // TS 15.05.15 umbau auf mehrere objekte
        //private void Cut() { Cut(DataAdapter.Instance.DataCache.Objects(Workspace).Object_Selected); }
        private void Cut() { Cut(DataAdapter.Instance.DataCache.Objects(Workspace).Objects_Selected); }
        public void Cut_Hotkey() { Cut(DataAdapter.Instance.DataCache.Objects(Workspace).Objects_Selected); }

        /// <summary>
        /// CServer setzt .ParentToCut und .isClipboardObject, sonst geschieht nichts
        /// </summary>
        /// <param name="obj"></param>
        // TS 15.05.15 umbau auf mehrere objekte
        //public void Cut(IDocumentOrFolder obj)
        public void Cut(List<IDocumentOrFolder> objectlist)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (_CanCut(objectlist))
                {
                    List<cmisObjectType> cutcmisobjects = new List<cmisObjectType>();
                    foreach (IDocumentOrFolder obj in objectlist)
                    {
                        IDocumentOrFolder objtest = GetValidCutObject(obj);
                        cutcmisobjects.Add(objtest.CMISObject);
                    }
                    DataAdapter.Instance.DataProvider.Cut(cutcmisobjects, DataAdapter.Instance.DataCache.Rights.UserPrincipal, new CallbackAction(Cut_Done));
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        // TS 15.05.15 umbau auf mehrere objekte
        public bool CanCut { get { return _CanCut(DataAdapter.Instance.DataCache.Objects(Workspace).Objects_Selected); } }

        private bool _CanCut()
        {
            return _CanCut(DataAdapter.Instance.DataCache.Objects(Workspace).Objects_Selected);
        }

        private bool _CanCut(List<IDocumentOrFolder> objectlist)
        {
            bool cancut = DataAdapter.Instance.DataCache.ApplicationFullyInit;
            // TS 12.11.13 auf original umbiegen falls dok und nicht original
            if (cancut)
            {
                foreach (IDocumentOrFolder obj in objectlist)
                {
                    IDocumentOrFolder objtest = GetValidCutObject(obj);
                    // TS 02.07.13 !obj.isClipboardObject dazu
                    if (cancut) cancut = objtest.canMoveObject
                        && !DataAdapter.Instance.DataCache.Objects(Workspace).IsDescendantCut(objtest.objectId)
                        && !DataAdapter.Instance.DataCache.Objects(Workspace).HasDescendantCut(objtest.objectId)
                        && !objtest.isClipboardObject;
                }
            }
            return cancut && objectlist != null && objectlist.Count > 0;
        }

        private IDocumentOrFolder GetValidCutObject(IDocumentOrFolder obj)
        {
            IDocumentOrFolder objtest = obj;
            if (obj.isDocument && obj.RefId != null && obj.RefId.Length > 0 &&
                (obj.CopyType.ToString().Equals(CSEnumDocumentCopyTypes.COPYTYPE_CONTENT.ToString())
                || obj.CopyType.ToString().Equals(CSEnumDocumentCopyTypes.COPYTYPE_TECHNICAL.ToString())
                ))
            {
                IDocumentOrFolder refparent = DataAdapter.Instance.DataCache.Objects(this.Workspace).GetObjectById(obj.RefId);
                while (
                    refparent != null &&
                    refparent.isDocument &&
                    refparent.RefId != null &&
                    refparent.RefId.Length > 0 &&
                    !refparent.RefId.Equals("0") &&
                    !refparent.CopyType.ToString().Equals(CSEnumDocumentCopyTypes.COPYTYPE_ORIGINAL.ToString()))
                {
                    refparent = DataAdapter.Instance.DataCache.Objects(this.Workspace).GetObjectById(refparent.RefId);
                }

                if (refparent != null && refparent.isDocument && refparent.CopyType.ToString().Equals(CSEnumDocumentCopyTypes.COPYTYPE_ORIGINAL.ToString()))
                    objtest = refparent;
            }
            return objtest;
        }

        /// <summary>
        /// info an alle
        /// </summary>
        private void Cut_Done()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            DataAdapter.Instance.InformObservers();
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        #endregion Cut (cmd_Cut)

        // ==================================================================

        #region Delete (cmd_Delete)

        private bool _deleteInProgress = false;
        public DelegateCommand cmd_Delete { get { return new DelegateCommand(Delete, _CanDelete); } }

        public void Delete()
        {
            Delete(null, null);
        }

        public void Delete(List<IDocumentOrFolder> objects, CallbackAction callback)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                bool highLevelObj = false;
                bool adressAkte = false;
                bool adressPerson = false;
                //string questiontext = Localization.localstring.msg_RequestForDelete;
                string questiontext = LocalizationMapper.Instance["msg_RequestForDelete"];

                if (_CanDelete(objects))
                {
                    // TS 04.02.15 gelöschte mails in papierkorb wenn sie nicht schon drin sind
                    List<cmisObjectType> cmisobjects = new List<cmisObjectType>();
                    if (objects == null)
                    {
                        objects = new List<IDocumentOrFolder>();
                        if (DataAdapter.Instance.DataCache.Objects(this.Workspace).Objects_Selected.Count > 0)
                        {
                            foreach (IDocumentOrFolder tmp in DataAdapter.Instance.DataCache.Objects(this.Workspace).Objects_Selected)
                                objects.Add(tmp);
                        }
                        else
                        {
                            IDocumentOrFolder selectedobject = DataAdapter.Instance.DataCache.Objects(this.Workspace).Object_Selected;
                            objects.Add(selectedobject);
                        }
                    }

                    foreach (IDocumentOrFolder tmp in objects)
                    {
                        // TS 04.02.15 gelöschte mails in papierkorb wenn sie nicht schon drin sind
                        cmisobjects.Add(tmp.CMISObject);

                        // Check for high level objects
                        highLevelObj = tmp.structLevel <= 8 && tmp.hasChildObjects ? true : highLevelObj;

                        // TS 18.01.16 bei adressen auch personen abfragen wenn keine kinder vorhanden sind
                        adressAkte = tmp.isAdressenAkte;
                        adressPerson = tmp.isAdressenVorgang;
                        highLevelObj = highLevelObj || adressAkte || adressPerson;
                    }

                    // Show dialog for accepting the delete on high level objects
                    if (cmisobjects.Count > 0)
                    {
                        bool saveIt = true;
                        int optionAskForDelete = DataAdapter.Instance.DataCache.Profile.Option_GetValueInteger(CSEnumOptions.alwaysaskondelete);
                        if (optionAskForDelete == Constants.PROFILE_OPTION_ASKONDELETE_ALWAYS || (optionAskForDelete == Constants.PROFILE_OPTION_ASKONDELETE_ONLY_ON_PARENTS && highLevelObj))
                        {
                            // TS 18.01.16 fallunterscheidung
                            //saveIt = showYesNoDialog(Localization.localstring.msg_RequestForDelete);
                            if (cmisobjects.Count == 1 && adressAkte) questiontext = LocalizationMapper.Instance["msg_RequestForDeleteFirma"];
                            if (cmisobjects.Count == 1 && adressPerson) questiontext = LocalizationMapper.Instance["msg_RequestForDeletePerson"];
                            saveIt = showYesNoDialog(questiontext);
                        }

                        if (saveIt)
                        {
                            _deleteInProgress = true;
                            DataAdapter.Instance.DataProvider.Delete(cmisobjects, DataAdapter.Instance.DataCache.Rights.UserPrincipal, new CallbackAction(Delete_Done, objects, callback));
                        }
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        /// <summary>
        /// Checks the list of given objects for any kind of linked relation-objects
        /// Then deletes the relations on both sides if possible (only in local datacache)
        /// </summary>
        /// <param name="sourceObjects"></param>
        private void DeleteRelationsForObjects(List<IDocumentOrFolder> sourceObjects)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                foreach (IDocumentOrFolder srcObj in sourceObjects)
                {
                    if (srcObj.hasRelationships)
                    {
                        List<string> targetIDs = new List<string>();
                        List<IDocumentOrFolder> targetObjects = new List<IDocumentOrFolder>();
                        targetIDs.AddRange(srcObj.getRelationshipSources);

                        foreach (string relID in targetIDs)
                        {
                            targetObjects.AddRange(DataAdapter.Instance.DataCache.Object_FindObjectWithID(relID));
                        }

                        foreach (IDocumentOrFolder obj in targetObjects)
                        {
                            obj.removeRelationship(srcObj.objectId);
                            if (ViewManager.IsComponentVisible(DataAdapter.Instance.DataCache.Profile.Profile_GetComponentFromLayout(CSEnumProfileComponentType.EDESKTOP, Constants.PROFILE_COMPONENTID_EDESKTOP)))
                            {
                                obj.NotifyPropertyChanged(CSEnumCmisProperties.cmis_lastModificationDate.ToString());
                            }
                        }
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        public bool CanDelete { get { return _CanDelete(null); } }

        private bool _CanDelete()
        {
            return _CanDelete(null);
        }

        public bool _CanDelete(List<IDocumentOrFolder> objects)
        {
            bool candelete = DataAdapter.Instance.DataCache.ApplicationFullyInit && !_deleteInProgress;
            bool isJobPending = false;
            bool isDeleteRequestActive = DataAdapter.Instance.DataCache.Profile.Option_GetValueBoolean(CSEnumOptions.enabledeleterequests);

            try
            {
                candelete = candelete && !isDeleteRequestActive;
                if (candelete)
                {
                    if (objects == null)
                    {
                        objects = new List<IDocumentOrFolder>();
                        if (DataAdapter.Instance.DataCache.Objects(this.Workspace).Objects_Selected.Count > 0)
                        {
                            foreach (IDocumentOrFolder tmp in DataAdapter.Instance.DataCache.Objects(this.Workspace).Objects_Selected)
                                objects.Add(tmp);
                        }
                        else
                        {
                            IDocumentOrFolder selectedobject = DataAdapter.Instance.DataCache.Objects(this.Workspace).Object_Selected;
                            objects.Add(selectedobject);
                        }
                    }
                    candelete = false;
                    if (objects.Count > 0)
                    {
                        candelete = true;
                        foreach (IDocumentOrFolder tmp in objects)
                        {
                            // Additional check for pending jobs of some types
                            isJobPending = WSPendingJobHelper.ExistsCriticalJobs(tmp.objectId);
                            
                            bool delPermission = false;
                            if (tmp.ACL != null)
                            {
                                delPermission = tmp.hasDeletePermission();
                            }
                            // Check if delete is allowed 
                            if (isJobPending || (!tmp.canDeleteTree && !tmp.canDeleteObject) || tmp.isAnnotationFile || (tmp.ACL != null && !delPermission))
                            {
                                candelete = false;
                                break;
                            }
                        }
                    }
                }
            }
            catch (System.Exception ex) { if (ex.Message.Length > 0) { Log.Log.Debug(ex.Message); } }
            return candelete;
        }

        private void Delete_Done(object deletedobjects, CallbackAction callback)
        {
            _deleteInProgress = false;
            if (DataAdapter.Instance.DataCache.ResponseStatus.success)
            {
                DeleteRelationsForObjects((List<IDocumentOrFolder>)deletedobjects);
                RefreshDZCollectionsAfterProcess(deletedobjects, null);

                if(callback != null)
                {
                    callback.Invoke();
                }
            }
        }

        #endregion Delete (cmd_Delete)

        public bool showYesNoDialog(string text)
        {
            // Create header-bar
            string appName = DataAdapter.Instance.DataCache.Profile.ProfileApplication.name;
            string appMandant = DataAdapter.Instance.DataCache.Rights.UserPrincipal.mandantname;
            string appVersion = Config.Config.versioninfo;
            char splitter = (".".ToCharArray())[0];
            string[] tokens = appVersion.Split(splitter);
            if (tokens != null && tokens.Length > 2)
            {
                // TS 18.01.16
                //appVersion = tokens[0] + "." + tokens[1];
                appVersion = tokens[0] + "." + tokens[1] + "." + tokens[2];
            }

            appName = appName.ToUpper().StartsWith("2CHARTA") ? appName : "2Charta " + appName;
            string boxHeader = appName + " " + appVersion + " " + appMandant;

            // show messagebox
            MessageBoxResult res = MessageBox.Show(text, boxHeader, MessageBoxButton.OKCancel);
            return (res == MessageBoxResult.OK);
        }

        public static void showOKDialog(string text)
        {
            // show messagebox
            MessageBoxResult res = showMessageBox(text, MessageBoxButton.OK);
        }

        public static bool showOKCancelDialog(string text)
        {
            // show messagebox
            MessageBoxResult res = showMessageBox(text, MessageBoxButton.OKCancel);
            return (res == MessageBoxResult.OK);
        }

        public static MessageBoxResult showMessageBox(string text, MessageBoxButton buttons)
        {
            return showMessageBox(null, text, buttons);
        }

        public static MessageBoxResult showMessageBox(Window window, string text, MessageBoxButton buttons)
        {
            // Create header-bar
            MessageBoxResult res;
            string appName = DataAdapter.Instance.DataCache.Profile.ProfileApplication.name;
            string appMandant = DataAdapter.Instance.DataCache.Rights.UserPrincipal.mandantname;
            string appVersion = Config.Config.versioninfo;
            char splitter = (".".ToCharArray())[0];
            string[] tokens = appVersion.Split(splitter);
            if (tokens != null && tokens.Length > 2)
                appVersion = tokens[0] + "." + tokens[1];
            appName = appName == null ? "2Charta ECM" : appName;
            appName = appName.ToUpper().StartsWith("2CHARTA") ? appName : "2Charta " + appName;
            string boxHeader = appName + " " + appVersion + " " + appMandant;

            // show messagebox            
            if(window != null)
            {
                res = MessageBox.Show(window, text, boxHeader, MessageBoxButton.OKCancel);
            }
            else
            {
                res = MessageBox.Show(text, boxHeader, MessageBoxButton.OKCancel);
            }            
            return res;
        }

        // ==================================================================

        #region DisplayExternal (cmd_DisplayExternal)

        public DelegateCommand cmd_DisplayExternal { get { return new DelegateCommand(DisplayExternal, _CanDisplayExternal); } }
        public DelegateCommand cmd_DisplayExternalAsPDF { get { return new DelegateCommand(DisplayExternalAsPDF, _CanDisplayExternalAsPDF); } }
        public DelegateCommand cmd_DisplayExternalAsHTML { get { return new DelegateCommand(DisplayExternalAsHTML, _CanDisplayExternalAsHTML); } }

        public void DisplayExternal_Hotkey() { DisplayExternal(); }
        public void DisplayExternal()
        {
            DisplayExternal(false, false);
        }

        public void DisplayExternalAsPDF()
        {
            DisplayExternal(true, false);
        }

        private bool _CanDisplayExternalAsPDF()
        {
            return _CanDisplayExternal() && DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected.DownloadNamePDF != null && DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected.DownloadNamePDF.Length > 0;
        }

        public void DisplayExternalAsHTML()
        {
            DisplayExternal(false, true);
        }

        private bool _CanDisplayExternalAsHTML()
        {
            return _CanDisplayExternal() && DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected.DownloadNameHTML != null && DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected.DownloadNameHTML.Length > 0;
        }

        private void DisplayExternal(bool forcepdf, bool forcehtml)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanDisplayExternal)
                {
                    string downloadname = DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected.DownloadName;
                    // TS 12.09.14 für mails die pdf dateien verwenden
                    // TS 15.05.15 auch für richtext documente und auch thml für den fehlerfall
                    if (DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected.isMailDocument || DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected.isRichTextDoc)
                    {
                        string downloadnamepdf = DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected.DownloadNamePDF;
                        if (downloadnamepdf != null && downloadnamepdf.Length > 0)
                        {
                            downloadname = downloadnamepdf;
                        }
                        else
                        {
                            string downloadnamehtml = DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected.DownloadNameHTML;
                            if (downloadnamehtml != null && downloadnamehtml.Length > 0)
                            {
                                downloadname = downloadnamehtml;
                            }
                        }
                    }

                    // TS 28.04.16
                    if (forcepdf)
                    {
                        string downloadnamepdf = DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected.DownloadNamePDF;
                        if (downloadnamepdf != null && downloadnamepdf.Length > 0)
                        {
                            downloadname = downloadnamepdf;
                        }
                    }
                    if (forcehtml)
                    {
                        string downloadnamehtml = DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected.DownloadNameHTML;
                        if (downloadnamehtml != null && downloadnamehtml.Length > 0)
                        {
                            downloadname = downloadnamehtml;
                        }
                    }

                    // TS 01.07.15 immer in echtem externen fenster anzeigen wie zuletzt nur für bpm
                    //if (downloadname != null && downloadname.Length > 0)
                    //{
                    //    System.Windows.Browser.HtmlWindow newWindow;
                    //    newWindow = HtmlPage.Window.Navigate(new Uri(downloadname, UriKind.RelativeOrAbsolute), "_blank", "");
                    //    //HtmlPage.Window.Navigate(new Uri(downloadname, UriKind.RelativeOrAbsolute), "_blank", "");
                    //}
                    //else
                    //{
                    //    // TS 23.01.15
                    //    downloadname = DataAdapter.Instance.DataCache.Objects(Workspace).Folder_Selected.DownloadName;
                    //    if (downloadname != null && downloadname.Length > 0)
                    //    {
                    //        System.Windows.Browser.HtmlWindow newWindow;
                    //        //newWindow = HtmlPage.Window.Navigate(new Uri(downloadname, UriKind.RelativeOrAbsolute), "_blank", "");
                    //        newWindow = HtmlPage.Window.Navigate(new Uri(downloadname, UriKind.RelativeOrAbsolute), "_blank", "width=1400, height=900,resizable=1,scrollbars=1,left=5,top=40,alwaysRaised,location=no");

                    //        //window.open(ur', 'popup', 'width=1775, height=950,resizable=1,scrollbars=1,left=5,top=40,alwaysRaised,location=no')

                    //    }
                    //}
                    // TS 01.07.15 umbau
                    // für bpm auf folderebene suchen
                    bool isbpm = false;
                    if (downloadname == null || downloadname.Length == 0)
                    {
                        IDocumentOrFolder selObj = DataAdapter.Instance.DataCache.Objects(Workspace).Folder_Selected;
                        string workFlowType = selObj.GetCMISPropertyAllValues(CSEnumCmisProperties.BPM_WORKFLOW).First();

                        isbpm = true;
                        downloadname = selObj.DownloadName;

                        //Show it in the seperated frame
                        string workflowSmallURL = selObj.GetCMISPropertyAllValues(CSEnumCmisProperties.BPM_SMALLMODE_URL).First();
                        if (workflowSmallURL != null && workflowSmallURL.Length > 0)
                        {
                            BPM_DisplaySideFrame(null, null, workflowSmallURL);
                            string workflowID = selObj.GetCMISPropertyAllValues(CSEnumCmisProperties.BPM_WORKFLOWID).First();

                            //Get Child objects
                            DataAdapter.Instance.DataProvider.BPM_GetDocuments(workflowID, CSEnumProfileWorkspace.workspace_workflow, DataAdapter.Instance.DataCache.Rights.UserPrincipal);
                        }
                        //
                    }
                    if (!isbpm && downloadname != null && downloadname.Length > 0)
                    {
                        DisplayExternal(downloadname);
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public bool CanDisplayExternal { get { return _CanDisplayExternal(); } }

        // TS 08.07.15
        public void DisplayExternal(string url)
        {
            System.Windows.Browser.HtmlWindow newWindow;
            if(url.ToUpper().StartsWith("WWW."))
            {
                url = "http://" + url;
            }

            bool newwindow = DataAdapter.Instance.DataCache.Profile.Option_GetValueBoolean(CSEnumOptions.urlexternalnewwindow);

            if (newwindow)
            {
                // anzeige in eigenem browserfenster
                newWindow = HtmlPage.Window.Navigate(new Uri(url, UriKind.Absolute), "_blank", "width=1400, height=900,resizable=1,scrollbars=1,left=5,top=40,alwaysRaised,location=no");
            }
            else
            {
                // anzeige im tab
                newWindow = HtmlPage.Window.Navigate(new Uri(url, UriKind.Absolute), "_blank", "");
            }
        }

        // TS 23.01.15 auch für ordner wg. bpm links
        //private bool _CanDisplayExternal() { return DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected.DownloadName.Length > 0; }
        private bool _CanDisplayExternal()
        {
            // TS 27.01.15 so klappen keine dokumente mehr
            //return DataAdapter.Instance.DataCache.Objects(Workspace).Object_Selected.DownloadName.Length > 0 || DataAdapter.Instance.DataCache.Objects(Workspace).Folder_Selected.DownloadName.Length > 0;
            return DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected.DownloadName.Length > 0 || DataAdapter.Instance.DataCache.Objects(Workspace).Folder_Selected.DownloadName.Length > 0;
        }

        #endregion DisplayExternal (cmd_DisplayExternal)

        // ==================================================================

        #region EditDefaults (cmd_EditDefaults)

        public DelegateCommand cmd_EditDefaults { get { return new DelegateCommand(EditDefaults, _CanEditDefaults); } }

        public void EditDefaults()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanEditDefaults)
                {
                    SL2C_Client.View.Dialogs.dlgEditDefaults child = new SL2C_Client.View.Dialogs.dlgEditDefaults(this.Workspace, false);
                    DialogHandler.Show_Dialog(child);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        public bool CanEditDefaults { get { return _CanEditDefaults(); } }

        private bool _CanEditDefaults()
        {
            return DataAdapter.Instance.DataCache.ApplicationFullyInit;
        }

        #endregion EditDefaults (cmd_EditDefaults)

        // ==================================================================

        #region EditQueryDefaults (cmd_EditQueryDefaults)

        public DelegateCommand cmd_EditQueryDefaults { get { return new DelegateCommand(EditQueryDefaults, _CanEditQueryDefaults); } }

        public void EditQueryDefaults()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanEditQueryDefaults)
                {
                    SL2C_Client.View.Dialogs.dlgEditDefaults child = new SL2C_Client.View.Dialogs.dlgEditDefaults(this.Workspace, true);
                    DialogHandler.Show_Dialog(child);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        public bool CanEditQueryDefaults { get { return _CanEditQueryDefaults(); } }

        private bool _CanEditQueryDefaults()
        {
            return DataAdapter.Instance.DataCache.ApplicationFullyInit;
        }

        #endregion EditQueryDefaults (cmd_EditQueryDefaults)

        // ==================================================================

        #region EditMeta (cmd_EditMeta)

        public DelegateCommand cmd_EditMeta { get { return new DelegateCommand(EditMeta, _CanEditMeta); } }

        public void EditMeta_Hotkey() { EditMeta(); }
        public void EditMeta()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanEditMeta)
                {
                    // TS 02.12.14 editmode aktivieren
                    DataAdapter.Instance.DataCache.Objects(Workspace).ObjectId_EditMode = DataAdapter.Instance.DataCache.Objects(Workspace).Folder_Selected.objectId;
                    SL2C_Client.View.Dialogs.dlgEditMeta child = new SL2C_Client.View.Dialogs.dlgEditMeta(this.Workspace);
                    DialogHandler.Show_Dialog(child);

                    //SL2C_Client.View.Dialogs.dlgEditDefaults child = new SL2C_Client.View.Dialogs.dlgEditDefaults(this.Workspace, true, false);
                    //child.Show();
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        public bool CanEditMeta { get { return _CanEditMeta(); } }

        private bool _CanEditMeta()
        {
            // TS 05.06.14 das dazu: isEditEnabled_ADR
            // TS 02.12.14 umgebaut wg. editmode
            // || DataAdapter.Instance.DataCache.Objects(this.Workspace).Folder_Selected.isEditEnabled
            // || DataAdapter.Instance.DataCache.Objects(this.Workspace).Folder_Selected.isEditEnabled_ADR
            // TS 14.01.15
            //return DataAdapter.Instance.DataCache.ApplicationFullyInit
            //    &&
            //    (DataAdapter.Instance.DataCache.Objects(this.Workspace).Folder_Selected.isNotCreated
            //    ||
            //    DataAdapter.Instance.DataCache.Objects(this.Workspace).Folder_Selected.isInWork
            //    ||
            //    DataAdapter.Instance.DataCache.Objects(this.Workspace).Folder_Selected.isEditEnabled
            //    ||
            //    DataAdapter.Instance.DataCache.Objects(this.Workspace).Folder_Selected.isEditEnabled_ADR
            //    );
            //return CanEditMode();
            // TS 26.05.15 prüfung auf objektid dazu
            return DataAdapter.Instance.DataCache.Objects(this.Workspace).Folder_Selected.objectId.Length > 0 && DataAdapter.Instance.DataCache.Objects(this.Workspace).Folder_Selected.isEditPossible;
        }

        #endregion EditMeta (cmd_EditMeta)

        // ==================================================================

        #region EditMode (cmd_EditMode)

        public DelegateCommand cmd_EditMode { get { return new DelegateCommand(EditMode, CanEditMode); } }

        public void EditMode_Hotkey() { EditMode(); }

        public void EditMode()
        {
            EditMode(false);
        }

        public void EditMode(bool isaddress)
        {
            if (CanEditMode(isaddress))
            {
                if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
                try
                {
                    DataAdapter.Instance.DataCache.Objects(Workspace).ObjectId_EditMode = DataAdapter.Instance.DataCache.Objects(Workspace).Folder_Selected.objectId;

                    NotifyPropertyChanged(null);
                    DataAdapter.Instance.InformObservers(this.Workspace);
                }
                catch (Exception e) { Log.Log.Error(e); }
                if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
            }
        }

        private bool CanEditMode()
        {
            return CanEditMode(false);
        }

        public bool CanEditMode(bool isaddress)
        {
            bool ret = false;
            // TS 13.01.15 autoedit prüfung dazu
            bool isautoedit = DataAdapter.Instance.DataCache.Profile.Option_GetValueBoolean(CSEnumOptions.automode_edit);
            IDocumentOrFolder dummy = DataAdapter.Instance.DataCache.Objects(Workspace).Folder_Selected;

            // TS 03.07.15 dämliche adress unterscheidung leider auch hier
            if (isaddress)
                ret = dummy.isAddressObject;
            else
            {
                ret = dummy.isPureOfficeObject;
            }

            if (ret)
            {
                ret = (DataAdapter.Instance.DataCache.ApplicationFullyInit && !isautoedit
                    && !dummy.isNotCreated && !dummy.isEdited && !DataAdapter.Instance.DataCache.Objects(Workspace).ObjectId_EditMode.Equals(dummy.objectId) && (dummy.canUpdateProperties || dummy.canSetContentStream));
            }
            return ret;
        }

        #endregion EditMode (cmd_EditMode)

        // ==================================================================

        #region EditObject

        public void EditObject(string objectid)
        {
            if (CanEditObject(objectid))
            {
                if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
                try
                {
                    IDocumentOrFolder dummy = DataAdapter.Instance.DataCache.Objects(Workspace).GetObjectById(objectid);
                    List<IDocumentOrFolder> changed = new List<IDocumentOrFolder>();
                    changed.Add(dummy);

                    // TS 15.05.14 validate zuerst
                    // TS 01.06.18 rausgenommen da vermutlich nicht mehr benoetigt weil hier nicht mehr gespeichert wird 
                    // und es sonst beim ersten fuellen eines feldes sofort das erste pflichtfeld anmeckert und das soll weg (#203-2 testliste)
                    //bool isvalid = dummy.ValidateData();

                    //if (DataAdapter.Instance.DataCache.Profile.Option_GetValueBoolean(CSEnumOptions.autosave_edit) && isvalid)
                    //{
                    //    Save(changed, null);
                    //}
                    //else
                    //{
                    ((IDataObserver)DataAdapter.Instance).ApplyData(changed);
                    // TS 02.02.15 kann raus da applydata bereits den inform macht
                    //DataAdapter.Instance.InformObservers();
                    //}
                }
                catch (Exception e) { Log.Log.Error(e); }
                if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
            }
        }

        public bool CanEditObject(string objectid)
        {
            return (DataAdapter.Instance.DataCache.ApplicationFullyInit
                && (
                DataAdapter.Instance.DataCache.Objects(Workspace).GetObjectById(objectid).isNotCreated
                || DataAdapter.Instance.DataCache.Objects(Workspace).GetObjectById(objectid).canUpdateProperties
                || DataAdapter.Instance.DataCache.Objects(Workspace).GetObjectById(objectid).canSetContentStream)
                );
        }

        #endregion EditObject

        // ==================================================================

        #region EnableEditDMSOffice_Level7

        public DelegateCommand cmd_EnableEditDMSOffice_Level7 { get { return new DelegateCommand(EnableEditDMSOffice_Level7, CanEnableEditDMSOffice_Level7); } }

        public void EnableEditDMSOffice_Level7()
        {
            if (CanEnableEditDMSOffice_Level7())
            {
                if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
                try
                {
                    EditMode();
                    IDocumentOrFolder dummy = DataAdapter.Instance.DataCache.Objects(Workspace).Folder_Selected;

                    string matchcode = dummy[Constants.IDX_KIL_MATCHCODE_NEWMODEL];
                    string firma1 = dummy[Constants.IDX_ADR_FIRMA_NAME1_NEWMODEL];
                    if (dummy.isOldADRModel)
                    {
                        matchcode = dummy[Constants.IDX_KIL_MATCHCODE_OLDMODEL];
                        firma1 = dummy[Constants.IDX_ADR_FIRMA_NAME1_OLDMODEL];

                        if (matchcode == null || matchcode.Length == 0)
                        {
                            if (firma1 != null && firma1.Length > 0)
                                dummy[Constants.IDX_KIL_MATCHCODE_OLDMODEL] = firma1;
                            else
                                dummy[Constants.IDX_KIL_MATCHCODE_OLDMODEL] = " ";
                        }
                    }
                    else
                    {
                        if (matchcode == null || matchcode.Length == 0)
                        {
                            if (firma1 != null && firma1.Length > 0)
                                dummy[Constants.IDX_KIL_MATCHCODE_NEWMODEL] = firma1;
                            else
                                dummy[Constants.IDX_KIL_MATCHCODE_NEWMODEL] = " ";
                        }
                    }

                    List<IDocumentOrFolder> changed = new List<IDocumentOrFolder>();
                    changed.Add(dummy);

                    ((IDataObserver)DataAdapter.Instance).ApplyData(changed);
                    DataAdapter.Instance.InformObservers();
                }
                catch (Exception e) { Log.Log.Error(e); }
                if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
            }
        }

        public bool CanEnableEditDMSOffice_Level7()
        {
            IDocumentOrFolder dummy = DataAdapter.Instance.DataCache.Objects(Workspace).Folder_Selected;
            return (DataAdapter.Instance.DataCache.ApplicationFullyInit
                && dummy.structLevel == Constants.STRUCTLEVEL_07_AKTE && dummy.isAdressenAkte && !dummy.isOfficeAkte
                && (dummy.isNotCreated || dummy.canUpdateProperties));
        }

        #endregion EnableEditDMSOffice_Level7

        // ==================================================================

        #region EnableEditADR_Level7

        public DelegateCommand cmd_EnableEditADR_Level7 { get { return new DelegateCommand(EnableEditADR_Level7, CanEnableEditADR_Level7); } }

        public void EnableEditADR_Level7()
        {
            if (CanEnableEditADR_Level7())
            {
                if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
                try
                {
                    EditMode();
                    IDocumentOrFolder dummy = DataAdapter.Instance.DataCache.Objects(Workspace).Folder_Selected;

                    string matchcode = dummy[Constants.IDX_KIL_MATCHCODE_NEWMODEL];
                    string firma1 = dummy[Constants.IDX_ADR_FIRMA_NAME1_NEWMODEL];
                    if (dummy.isOldADRModel)
                    {
                        matchcode = dummy[Constants.IDX_KIL_MATCHCODE_OLDMODEL];
                        firma1 = dummy[Constants.IDX_ADR_FIRMA_NAME1_OLDMODEL];
                        if (firma1 == null || firma1.Length == 0)
                        {
                            if (matchcode != null && matchcode.Length > 0)
                                dummy[Constants.IDX_ADR_FIRMA_NAME1_OLDMODEL] = matchcode;
                            else
                                dummy[Constants.IDX_ADR_FIRMA_NAME1_OLDMODEL] = " ";
                        }
                    }
                    else
                    {
                        if (firma1 == null || firma1.Length == 0)
                        {
                            if (matchcode != null && matchcode.Length > 0)
                                dummy[Constants.IDX_ADR_FIRMA_NAME1_NEWMODEL] = matchcode;
                            else
                                dummy[Constants.IDX_ADR_FIRMA_NAME1_NEWMODEL] = " ";
                        }
                    }


                    List<IDocumentOrFolder> changed = new List<IDocumentOrFolder>();
                    changed.Add(dummy);

                    ((IDataObserver)DataAdapter.Instance).ApplyData(changed);
                    DataAdapter.Instance.InformObservers();
                }
                catch (Exception e) { Log.Log.Error(e); }
                if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
            }
        }

        public bool CanEnableEditADR_Level7()
        {
            IDocumentOrFolder dummy = DataAdapter.Instance.DataCache.Objects(Workspace).Folder_Selected;
            return (DataAdapter.Instance.DataCache.ApplicationFullyInit
                && dummy.structLevel == Constants.STRUCTLEVEL_07_AKTE && !dummy.isAdressenAkte && dummy.isOfficeAkte
                && (dummy.isNotCreated || dummy.canUpdateProperties));
        }

        #endregion EnableEditADR_Level7

        // ==================================================================

        #region Export (cmd_Export)

        public DelegateCommand cmd_Export { get { return new DelegateCommand(Export, _CanExport); } }

        public DelegateCommand cmd_ExportNoMeta { get { return new DelegateCommand(ExportNoMeta, _CanExport); } }

        public void Export()
        {
            // TS 23.06.15 überarbeitet um auch aus der ablagemappe ungespeicherte (kopierte) verarbeiten zu können
            // und um die dateiversionen hier und nicht im Silverlight zu ermitteln
            //
            //														deep	paginate	input
            //Export			=> vollständiger rekursiver Export	true	true		all levels
            //
            //Print				=> nur Dateien und Dokumente		false	false		> = 9
            //
            //MergeFiles    	=> wie print + automatisch ablegen	false	false		> = 9
            //
            //{ConvertToPDF]	=> CServer intern, nach scannen		false	false		> = 10
            //
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanExport)
                {
                    Export(DataAdapter.Instance.DataCache.Objects(this.Workspace).Objects_Selected, true, true, true, new CallbackAction(Export_Done));
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public void ExportNoMeta()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanExport)
                {
                    Export(DataAdapter.Instance.DataCache.Objects(this.Workspace).Objects_Selected, true, true, false, new CallbackAction(Export_Done));
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public void Export(List<IDocumentOrFolder> objects, bool paginate, bool deep, bool addmeta, CallbackAction callback)
        {
            // TS 23.06.15 überarbeitet um auch aus der ablagemappe ungespeicherte (kopierte) verarbeiten zu können
            // und um die dateiversionen hier und nicht im Silverlight zu ermitteln
            //
            //														deep	paginate	input
            //Export			=> vollständiger rekursiver Export	true	true		all levels
            //
            //Print				=> nur Dateien und Dokumente		false	false		> = 9
            //
            //MergeFiles    	=> wie print + automatisch ablegen	false	false		> = 9
            //
            //{ConvertToPDF]	=> CServer intern, nach scannen		false	false		> = 10
            //
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (_CanExport(objects))
                {
                    List<cmisObjectType> exportobjects = new List<cmisObjectType>();
                    foreach (IDocumentOrFolder obj in objects)
                    {
                        exportobjects.Add(obj.CMISObject);
                    }

                    if (callback == null)
                        callback = new CallbackAction(Export_Done);
                    CSOption showcopytypeoption = DataAdapter.Instance.DataCache.Profile.Option_GetOption(CServer.CSEnumOptions.showcopytypes);
                    CSEnumOptionPresets_showcopytypes showcopytype = DataAdapter.Instance.DataCache.Profile.Option_GetEnumValue<CSEnumOptionPresets_showcopytypes>(showcopytypeoption);

                    // TS 30.11.17 sortierung aus treeview holen
                    string orderby = "";
                    List<CServer.CSOrderToken> givenorderby = new List<CSOrderToken>();
                    _Query_GetSortDataFilter(ref givenorderby, false);
                    foreach (CServer.CSOrderToken orderbytoken in givenorderby)
                    {
                        if (orderby.Length > 0)
                        {
                            // TS 01.10.13 muss komma sein statt blank !!
                            orderby = orderby + ", ";
                        }
                        orderby = orderby + orderbytoken.propertyname + " " + orderbytoken.orderby;
                    }

                    // TS 30.11.17 sortierung mitgeben
                    // DataAdapter.Instance.DataProvider.ExportPDF(exportobjects, paginate, deep, addmeta, showcopytype, DataAdapter.Instance.DataCache.Rights.UserPrincipal, callback);
                    DataAdapter.Instance.DataProvider.ExportPDF(exportobjects, paginate, deep, addmeta, showcopytype, orderby, DataAdapter.Instance.DataCache.Rights.UserPrincipal, callback);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        public bool CanExport { get { return _CanExport(DataAdapter.Instance.DataCache.Objects(Workspace).Objects_Selected); } }

        private bool _CanExport()
        {
            return _CanExport(DataAdapter.Instance.DataCache.Objects(Workspace).Objects_Selected);
        }

        private bool _CanExport(List<IDocumentOrFolder> objects)
        {
            //bool canexport = DataAdapter.Instance.DataCache.ApplicationFullyInit && obj.RepositoryId.Length > 0;
            bool canexport = DataAdapter.Instance.DataCache.ApplicationFullyInit && objects.Count > 0;
            return canexport;
        }

        private void Export_Done()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (DataAdapter.Instance.DataCache.ResponseStatus.success)
                {
                    string filenamezip = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.URL_ExportZIP];
                    string filenamepdf = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.URL_ExportPDF];
                    // extern anzeigen, dort kann dann gespeichert werden werden

                    // TS 21.12.15
                    //System.Windows.Browser.HtmlWindow newWindow;
                    //if (filenamezip != null && filenamezip.Length > 0)
                    //    newWindow = HtmlPage.Window.Navigate(new Uri(filenamezip, UriKind.RelativeOrAbsolute), "_blank", "");
                    //else if (filenamepdf != null && filenamepdf.Length > 0)
                    //    newWindow = HtmlPage.Window.Navigate(new Uri(filenamepdf, UriKind.RelativeOrAbsolute), "_blank", "");
                    if (filenamezip != null && filenamezip.Length > 0)
                        DisplayExternal(filenamezip);
                    else if (filenamepdf != null && filenamepdf.Length > 0)
                        DisplayExternal(filenamepdf);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        #endregion Export (cmd_Export)

        // ==================================================================

        #region ExportStruct (cmd_ExportStruct)

        public DelegateCommand cmd_ExportStruct { get { return new DelegateCommand(ExportStruct, _CanExportStruct); } }

        public void ExportStruct()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanExportStruct)
                {
                    ExportStruct(DataAdapter.Instance.DataCache.Objects(this.Workspace).Objects_Selected, new CallbackAction(ExportStruct_Done));
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public void ExportStruct(List<IDocumentOrFolder> objects, CallbackAction callback)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (_CanExportStruct(objects))
                {
                    List<cmisObjectType> exportobjects = new List<cmisObjectType>();
                    foreach (IDocumentOrFolder obj in objects)
                    {
                        exportobjects.Add(obj.CMISObject);
                    }

                    if (callback == null)
                        callback = new CallbackAction(ExportStruct_Done);
                    CSOption showcopytypeoption = DataAdapter.Instance.DataCache.Profile.Option_GetOption(CServer.CSEnumOptions.showcopytypes);
                    CSEnumOptionPresets_showcopytypes showcopytype = DataAdapter.Instance.DataCache.Profile.Option_GetEnumValue<CSEnumOptionPresets_showcopytypes>(showcopytypeoption);

                    DataAdapter.Instance.DataProvider.ExportStruct(exportobjects, showcopytype, DataAdapter.Instance.DataCache.Rights.UserPrincipal, callback);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        public bool CanExportStruct { get { return _CanExportStruct(DataAdapter.Instance.DataCache.Objects(Workspace).Objects_Selected); } }

        private bool _CanExportStruct()
        {
            return _CanExportStruct(DataAdapter.Instance.DataCache.Objects(Workspace).Objects_Selected);
        }

        public bool _CanExportStruct(List<IDocumentOrFolder> objects)
        {
            //bool canexport = DataAdapter.Instance.DataCache.ApplicationFullyInit && obj.RepositoryId.Length > 0;
            bool canexport = DataAdapter.Instance.DataCache.ApplicationFullyInit && objects.Count > 0;
            return canexport;
        }

        private void ExportStruct_Done()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (DataAdapter.Instance.DataCache.ResponseStatus.success)
                {
                    string filenamezip = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.URL_ExportZIP];
                    if (filenamezip != null && filenamezip.Length > 0)
                        DisplayExternal(filenamezip);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        #endregion Export (cmd_Export)


        // ==================================================================

        #region ExportCSV (cmd_ExportCSV)

        public DelegateCommand cmd_ExportCSV { get { return new DelegateCommand(ExportCSV, _CanExportCSV); } }

        public void ExportCSV()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanExportCSV)
                {
                    List<cmisObjectType> exportobjects = new List<cmisObjectType>();
                    foreach (IDocumentOrFolder obj in DataAdapter.Instance.DataCache.Objects(Workspace).ObjectList_QueryResults)
                    {
                        // TS 16.03.17 nur die ausgewählten übernehmen
                        // exportobjects.Add(obj.CMISObject);
                        if (obj.isUserSelected)
                        {
                            try
                            {
                                obj[CSEnumCmisProperties.tmp_RootNodeName] = obj.RootNodeName;
                            }
                            catch (Exception) { }
                            exportobjects.Add(obj.CMISObject);
                        }
                    }

                    SaveFileDialog dlg = _FileSaveDialog(null, "", "");
                    if (dlg != null && dlg.SafeFileName != null && dlg.SafeFileName.Length > 0)
                    {
                        // TS 21.11.16 keinen filter mehr aber dafür die aktuell gewählten spalten
                        //List<string> propsfilter = ExportCSV_GetPropertiesFilter();
                        List<string> propsfilter = new List<string>();
                        List<string> usedcolumns = ListDisplayProperties;
                        if (usedcolumns.Contains("RootNodeName"))
                        {
                            usedcolumns.Insert(usedcolumns.IndexOf("RootNodeName"), CSEnumCmisProperties.tmp_RootNodeName.ToString());
                            usedcolumns.Remove("RootNodeName");
                        }
                        if (usedcolumns.Contains("FULLTEXT"))
                        {
                            usedcolumns.Remove("FULLTEXT");
                        }
                        List<string> negfilter = new List<string>();
                        negfilter.Add("cmis:");

                        bool allvaluedimensions = false;
                        if (propsfilter.Contains("ADR_"))
                            allvaluedimensions = true;
                        bool useparentobjects = true;

                        // TS 21.11.16
                        //DataAdapter.Instance.DataProvider.ExportCSV(exportobjects, null, propsfilter, negfilter, allvaluedimensions, useparentobjects, DataAdapter.Instance.DataCache.Rights.UserPrincipal, new CallbackAction(ExportCSV_Done, (object)dlg));
                        // TS 28.06.17
                        //DataAdapter.Instance.DataProvider.ExportCSV(exportobjects, usedcolumns, propsfilter, negfilter, allvaluedimensions, useparentobjects, DataAdapter.Instance.DataCache.Rights.UserPrincipal, new CallbackAction(ExportCSV_Done, (object)dlg));
                        string appidcurrent = DataAdapter.Instance.DataCache.Rights.ChoosenApplication.id;
                        string repidcurrent = DataAdapter.Instance.DataCache.Rights.ChoosenRepository.repositoryId;
                        DataAdapter.Instance.DataProvider.ExportCSV(exportobjects, usedcolumns, propsfilter, negfilter, allvaluedimensions, useparentobjects, appidcurrent, repidcurrent, DataAdapter.Instance.DataCache.Rights.UserPrincipal, new CallbackAction(ExportCSV_Done, (object)dlg));
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        public bool CanExportCSV { get { return _CanExportCSV(); } }

        private bool _CanExportCSV()
        {
            bool canexport = DataAdapter.Instance.DataCache.ApplicationFullyInit && DataAdapter.Instance.DataCache.Objects(Workspace).ObjectList_QueryResults.Count() > 0;
            return canexport;
        }

        private List<string> ExportCSV_GetPropertiesFilter()
        {
            // je workspace / archiv:
            //KIL-default:
            //    KIL_
            //    DOKUMENT_
            //KIL_ADR (ISADR entscheidet)
            //    ADR_
            //    DOKUMENT_
            //Internal:
            //    INTERNAL_
            //    DOKUMENT_
            //Post:
            //    POST_
            //    DOKUMENT_
            //Mail:
            //    MAIL_
            //    DOKUMENT_

            string mandantid = DataAdapter.Instance.DataCache.Rights.UserPrincipal.mandantid;

            List<string> propsfilter = new List<string>();
            propsfilter.Add("DOKUMENT_");
            propsfilter.Add("DOKUMENTSYS_");
            switch (this.Workspace)
            {
                case CSEnumProfileWorkspace.workspace_default:
                    {
                        // TS 08.08.16 KIL nicht mehr vorhanden
                        //if (DataAdapter.Instance.DataCache.Repository(this.Workspace).RepositoryInfo.repositoryId.Equals(Constants.REPOSITORY_KIL + "_" + mandantid))
                        //{
                        IDocumentOrFolder current = DataAdapter.Instance.DataCache.Objects(this.Workspace).Folder_Selected;
                        // TS 09.08.16
                        if (current.isOldADRModel)
                        {
                            if (current.isAdressenAkte)
                            {
                                propsfilter.Add("KIL_");
                                propsfilter.Add("ADR_");
                            }
                            else if (current.isAdressenVorgang || current.isAdressenDokument)
                                propsfilter.Add("ADR_");
                            else
                                propsfilter.Add("KIL_");
                        }
                        else
                        {
                            if (current.isAdressenAkte)
                            {
                                propsfilter.Add("AKTE_");
                                propsfilter.Add("AKTESYS_");
                                propsfilter.Add("AKTE_ADR_");
                            }
                            else if (current.isAdressenVorgang)
                                propsfilter.Add("VORGANG_ADR_");
                            //else if (current.isAdressenDokument)
                            //    propsfilter.Add("DOKUMENT_");
                            else
                            {
                                propsfilter.Add("AKTE_");
                                propsfilter.Add("AKTESYS_");
                                propsfilter.Add("VORGANG_");
                                propsfilter.Add("VORGANGSYS_");
                                //propsfilter.Add("DOKUMENT_");
                            }
                        }
                    }
                    break;

                case CSEnumProfileWorkspace.workspace_post:
                    propsfilter.Add("POST_");
                    break;

                case CSEnumProfileWorkspace.workspace_mail:
                    propsfilter.Add("MAIL_");
                    break;

                case CSEnumProfileWorkspace.workspace_ablage:
                    propsfilter.Add("INTERNAL_");
                    break;

                case CSEnumProfileWorkspace.workspace_gesamt:
                    propsfilter.Add("INTERNAL_");
                    break;

                case CSEnumProfileWorkspace.workspace_stapel:
                    propsfilter.Add("INTERNAL_");
                    break;

                case CSEnumProfileWorkspace.workspace_aufgabe:
                    propsfilter.Add("INTERNAL_");
                    break;

                case CSEnumProfileWorkspace.workspace_aufgabepublic:
                    propsfilter.Add("INTERNAL_");
                    break;

                case CSEnumProfileWorkspace.workspace_termin:
                    propsfilter.Add("INTERNAL_");
                    break;
                    //case CSEnumProfileWorkspace.workspace_workflow:
                    //    propsfilter.Add("INTERNAL_");
                    //    break;
            }
            return propsfilter;
        }

        /// <summary>
        /// datei runterladen, SaveFileDialog kommt als Parameter mit
        /// </summary>
        private void ExportCSV_Done(object dlg)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (DataAdapter.Instance.DataCache.ResponseStatus.success && dlg.GetType().Name.Equals("SaveFileDialog"))
                {
                    string filenamezip = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.URL_ExportZIP];

                    SaveFileDialog dialog = (SaveFileDialog)dlg;
                    string downloadname = dialog.SafeFileName;
                    // TS 19.12.13
                    string filenamecsv = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.URL_ExportCSV];
                    new FileDownloadHandler().DownloadFile(filenamecsv, (Stream)dialog.OpenFile(), downloadname);
                    dialog = null;
                    dlg = null;
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        //private void ExportAsCSV_Done()
        //{
        //    if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
        //    try
        //    {
        //        if (DataAdapter.Instance.DataCache.ResponseStatus.success)
        //        {
        //            string filenamezip = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.URL_ExportAsZIP];
        //            string filenamepdf = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.URL_ExportAsPDF];
        //            // extern anzeigen, dort kann dann gespeichert werden werden
        //            System.Windows.Browser.HtmlWindow newWindow;
        //            if (filenamezip != null && filenamezip.Length > 0)
        //                newWindow = HtmlPage.Window.Navigate(new Uri(filenamezip, UriKind.RelativeOrAbsolute), "_blank", "");
        //            else if (filenamepdf != null && filenamepdf.Length > 0)
        //                newWindow = HtmlPage.Window.Navigate(new Uri(filenamepdf, UriKind.RelativeOrAbsolute), "_blank", "");
        //        }
        //    }
        //    catch (Exception e) { Log.Log.Error(e); }
        //    if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        //}

        #endregion ExportCSV (cmd_ExportCSV)

        // ==================================================================

        #region ExportFile (cmd_ExportFile)

        public DelegateCommand cmd_ExportFile { get { return new DelegateCommand(ExportFile, _CanExportFile); } }

        public void ExportFile()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanExportFile)
                {
                    IDocumentOrFolder selecteddoc = DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected;
                    string downloadname = selecteddoc.DownloadName;

                    // TS 17.07.17 wenn mail dann 3 verschiedene downloads anbieten: pdf, html und eml
                    // TS 17.07.17 fuer sonstige dateien immer nur den fixen zurgeundeliegenden dateityp
                    bool ismail = false;
                    string downloadname_pdf = "";
                    string downloadname_html = "";
                    string extension = "";
                    if (downloadname.Contains("."))
                    {
                        extension = downloadname.Substring(downloadname.LastIndexOf(".") + 1);
                    }
                    if (downloadname.ToLower().EndsWith(".mail"))
                    {
                        // TS 17.07.17
                        //string downloadnametmp = selecteddoc.DownloadNamePDF;
                        //if (downloadnametmp == null || downloadnametmp.Length == 0)
                        //    downloadnametmp = selecteddoc.DownloadNameHTML;
                        //if (downloadnametmp != null && downloadnametmp.Length > 0)
                        //    downloadname = downloadnametmp;
                        ismail = true;
                        downloadname = downloadname.Substring(0, downloadname.Length - 5);
                        downloadname_pdf = selecteddoc.DownloadNamePDF;
                        downloadname_html = selecteddoc.DownloadNameHTML;
                    }
                    string filename = downloadname;

                    if (filename.Contains("\\"))
                        filename = filename.Substring(filename.LastIndexOf("\\") + 1);
                    else if (filename.Contains("/"))
                        filename = filename.Substring(filename.LastIndexOf("/") + 1);

                    // TS 17.07.17
                    // SaveFileDialog dlg = _FileSaveDialog(null, filename, LocalizationMapper.Instance["dlgfilesavefilter_all"]);
                    string filter = LocalizationMapper.Instance["dlgfilesavefilter_all"];
                    if (ismail)
                    {
                        filter = "Standard-Mailformat (*.eml) | *.eml |PDF (*.pdf) | *.pdf |HTML (*.html) | *.html";
                    }
                    else if (extension.Length > 0)
                    {
                        filter = extension.ToUpper() + " (*." + extension + ") | *." + extension;
                    }                    

                    SaveFileDialog dlg = _FileSaveDialog(null, filename, filter);

                    if (dlg != null && dlg.SafeFileName != null && dlg.SafeFileName.Length > 0)
                    {
                        //SaveFileDialog dialog = (SaveFileDialog)dlg;
                        //string downloadname = dialog.SafeFileName;
                        // TS 17.07.17
                        if (ismail)
                        {
                            // die neue logik mit unterschiedlichen ausgabeformaten
                            switch (dlg.FilterIndex)
                            {
                                case 1:
                                    // info cache vorab leeren damit der neue EML downloadname dort abgelegt werden kann
                                    // DataAdapter.Instance.DataCache.Info[CSEnumInformationId.URL_ExportEML] = "";
                                    CSInformation[] clearinfoarray = new CSInformation[1];
                                    clearinfoarray[0] = new CSInformation();
                                    clearinfoarray[0].informationid = CSEnumInformationId.URL_ExportEML;
                                    clearinfoarray[0].informationidSpecified = true;
                                    clearinfoarray[0].informationvalue = "";
                                    DataAdapter.Instance.DataCache.Info.AddOrReplaceInformation(clearinfoarray);

                                    string repositoryid = selecteddoc.RepositoryId;
                                    string mailobjectid = selecteddoc.objectId;

                                    // ******************
                                    //IDocumentOrFolder currentdoc = DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected;
                                    //if (!currentdoc.isMailDocument)
                                    //{
                                    //    currentdoc = null;
                                    //    IDocumentOrFolder currentfld = DataAdapter.Instance.DataCache.Objects(Workspace).Folder_Selected;
                                    //    foreach (IDocumentOrFolder child in currentfld.ChildDocuments)
                                    //    {
                                    //        if (child.isMailDocument)
                                    //        {
                                    //            currentdoc = child;
                                    //            break;
                                    //        }
                                    //    }
                                    //}
                                    //if (currentdoc.isMailDocument)
                                    // ******************

                                    CallbackAction callback = new CallbackAction(FileSave_EMLCreated_Callback, dlg);
                                    DataAdapter.Instance.DataProvider.ExportEML(repositoryid, mailobjectid, DataAdapter.Instance.DataCache.Rights.UserPrincipal, callback);
                                    break;
                                case 2:
                                    new FileDownloadHandler().DownloadFile(downloadname_pdf, (Stream)dlg.OpenFile(), dlg.SafeFileName);
                                    break;
                                case 3:
                                    new FileDownloadHandler().DownloadFile(downloadname_html, (Stream)dlg.OpenFile(), dlg.SafeFileName);
                                    break;
                            }
                        }
                        else
                        {
                            // die alte logik
                            new FileDownloadHandler().DownloadFile(downloadname, (Stream)dlg.OpenFile(), dlg.SafeFileName);
                            dlg = null;
                        }
                        // TS 17.07.17 ins callback
                        //dlg = null;
                    }
                    //DataAdapter.Instance.DataProvider.ExportAsCSV(exportobjects, null, propsfilter, negfilter, true, false, DataAdapter.Instance.DataCache.Rights.UserPrincipal, new CallbackAction(ExportAsCSV_Done));
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        private void FileSave_EMLCreated_Callback(object savefiledialog)
        {
            if (DataAdapter.Instance.DataCache.ResponseStatus.success && savefiledialog.GetType().Name.Equals("SaveFileDialog") 
                && DataAdapter.Instance.DataCache.Info[CSEnumInformationId.URL_ExportEML] != null && DataAdapter.Instance.DataCache.Info[CSEnumInformationId.URL_ExportEML].Length > 0)
            {
                SaveFileDialog dlg = (SaveFileDialog)savefiledialog;
                string downloadname = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.URL_ExportEML];
                new FileDownloadHandler().DownloadFile(downloadname, (Stream)dlg.OpenFile(), dlg.SafeFileName);
                dlg = null;
            }
        }

        // ------------------------------------------------------------------
        public bool CanExportFile { get { return _CanExportFile(); } }

        private bool _CanExportFile()
        {
            // TS 17.07.17 auch erlauben wenn auf doklog stehend
            //IDocumentOrFolder selecteddoc = DataAdapter.Instance.DataCache.Objects(Workspace).Object_Selected;
            IDocumentOrFolder selecteddoc = DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected;
            bool canexport = selecteddoc.structLevel == Constants.STRUCTLEVEL_10_DOKUMENT;
            if(canexport) canexport = DataAdapter.Instance.DataCache.ApplicationFullyInit && selecteddoc.objectId.Length > 0;
            return canexport;
        }

        #endregion ExportFile (cmd_ExportFile)

        // ==================================================================

        #region ExportOffline (cmd_ExportOffline)

        public DelegateCommand cmd_ExportOffline { get { return new DelegateCommand(ExportOffline, _CanExportOffline); } }

        public void ExportOffline()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanExportOffline)
                {
                    List<cmisObjectType> exportObjectsCmis = new List<cmisObjectType>();
                    exportObjectsCmis.Add(DataAdapter.Instance.DataCache.Objects(this.Workspace).Object_Selected.CMISObject);

                    List<IDocumentOrFolder> exportObjects = new List<IDocumentOrFolder>();
                    exportObjects.Add(DataAdapter.Instance.DataCache.Objects(this.Workspace).Object_Selected);

                    SaveFileDialog dlg = _FileSaveDialog(null, "", LocalizationMapper.Instance["dlgfilesavefilter_zip"]);
                    if (dlg != null && dlg.SafeFileName != null && dlg.SafeFileName.Length > 0)
                    {
                        string repid = DataAdapter.Instance.DataCache.Rights.ChoosenRepository.repositoryId;
                        if (repid != null && repid.Length > 0 && repid.Contains("_"))
                        {
                            repid = repid.Substring(0, repid.IndexOf("_"));
                        }

                        // Start export
                        if (exportObjects[0].isNewDatamodel)
                        {
                            // Checkout is only activated in the new datamodel!
                            CallbackAction callback = new CallbackAction(ExportOffline_preFinalizer, dlg, exportObjects, exportObjectsCmis, repid);
                            SetCheckoutState(exportObjects, true, Constants.SPEC_PROPERTY_STATUS_CHECKOUT, callback);
                        }else
                        {
                            ExportOffline_preFinalizer(dlg, exportObjects, exportObjectsCmis, repid);
                        }
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        private void ExportOffline_preFinalizer(object dialog, object expObjects, object expObjectsCmis, object repidObj)
        {
            List<IDocumentOrFolder> exportObjects = (List<IDocumentOrFolder>)expObjects;
            List<cmisObjectType> exportObjectsCmis = (List<cmisObjectType>)expObjectsCmis;
            string repid = (string)repidObj;

            CallbackAction finalCallback = new CallbackAction(ExportOffline_Done, dialog, exportObjects);
            DataAdapter.Instance.DataProvider.ExportOffline(exportObjectsCmis, DataAdapter.Instance.DataCache.Rights.ChoosenApplication.id, repid, DataAdapter.Instance.DataCache.Rights.UserPrincipal, finalCallback);
        }

        // ------------------------------------------------------------------
        public bool CanExportOffline { get { return _CanExportOffline(); } }

        private bool _CanExportOffline()
        {
            bool canexport = DataAdapter.Instance.DataCache.ApplicationFullyInit
                && DataAdapter.Instance.DataCache.Objects(this.Workspace).Objects_Selected.Count() > 0
                && DataAdapter.Instance.DataCache.Objects(Workspace).Object_Selected.structLevel >= Statics.Constants.STRUCTLEVEL_07_AKTE
                && DataAdapter.Instance.DataCache.Objects(Workspace).Object_Selected.structLevel <= Statics.Constants.STRUCTLEVEL_09_DOKLOG
                &&
                (
                this.Workspace.ToString().Equals(CSEnumProfileWorkspace.workspace_default.ToString())
                ||
                this.Workspace.ToString().Equals(CSEnumProfileWorkspace.workspace_clipboard.ToString())
                );
            return canexport;
        }

        /// <summary>
        /// datei runterladen, SaveFileDialog kommt als Parameter mit
        /// </summary>
        private void ExportOffline_Done(object dlg, object exportedObjects)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (DataAdapter.Instance.DataCache.ResponseStatus.success && dlg.GetType().Name.Equals("SaveFileDialog"))
                {
                    SaveFileDialog dialog = (SaveFileDialog)dlg;
                    string downloadname = dialog.SafeFileName;
                    // TS 19.12.13
                    string filenamezip = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.URL_ExportOffline];
                    new FileDownloadHandler().DownloadFile(filenamezip, (Stream)dialog.OpenFile(), dialog.SafeFileName);
                    dialog = null;
                    dlg = null;
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        #endregion ExportOffline (cmd_ExportOffline)

        // ==================================================================

        // ==================================================================

        #region Fav_CreateFSLink (cmd_Fav_CreateFSLink)

        public DelegateCommand cmd_Fav_CreateFSLink { get { return new DelegateCommand(Fav_CreateFSLink, _CanFav_CreateFSLink); } }

        public void Fav_CreateFSLink() { Fav_CreateFSLink(null); }

        public void Fav_CreateFSLink(string filename)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (_CanFav_CreateFSLink())
                {
                    // spezielle asynchrone anbindung: der request wird an lc geschickt
                    // die antwort schreibt sich in den datacache_info und wird im rclistener abgefangen um dann den dialog zu rufen

                    if (filename == null || filename.Length == 0)
                    {
                        DataAdapter.Instance.DataCache.Info.RemoveInformation(CSEnumInformationId.Undefined);
                        CSInformation info = new CSInformation();
                        info.informationid = CSEnumInformationId.Undefined;
                        info.informationvalue = CSEnumCommands.cmd_Fav_CreateFSLink.ToString();
                        CSInformation[] infolist = new CSInformation[1];
                        infolist[0] = info;
                        DataAdapter.Instance.DataCache.Info.AddOrReplaceInformation(infolist);

                        this.LC_FileSelect("");
                    }
                    else
                    {
                        // wenn antwort erhalten dann
                        CSInformation info = new CSInformation();
                        info.informationid = CSEnumInformationId.Undefined;
                        info.informationvalue = filename;
                        CSInformation[] infolist = new CSInformation[1];
                        infolist[0] = info;
                        DataAdapter.Instance.DataCache.Info.AddOrReplaceInformation(infolist);

                        _CreateFavOrLink(CSEnumInternalObjectType.Favorit, CSEnumInternalObjectType.LinkFS, CSEnumProfileWorkspace.workspace_favoriten, Constants.STRUCTLEVEL_09_DOKLOG, null, true);
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        public bool CanFav_CreateFSLink { get { return _CanFav_CreateFSLink(); } }

        public bool _CanFav_CreateFSLink()
        {
            IDocumentOrFolder selObj = DataAdapter.Instance.DataCache.Objects(Workspace).Object_Selected;
            return _CanCreateFoldersInternal(CSEnumInternalObjectType.LinkFS, CSEnumProfileWorkspace.workspace_favoriten, Constants.STRUCTLEVEL_09_DOKLOG, null)
                && (selObj == null || selObj.objectId.Length == 0);
        }

        // ------------------------------------------------------------------

        #endregion Fav_CreateFSLink (cmd_Fav_CreateFSLink)        

        // ==================================================================

        #region Fav_CreateLastEditedLink (cmd_Fav_CreateLastEditedLink)

        public DelegateCommand cmd_Fav_CreateLastEditedLink { get { return new DelegateCommand(Fav_CreateLastEditedLink, _CanFav_CreateLastEditedLink); } }

        //public void Fav_CreateLastEditedLink()
        //{
        //    Fav_CreateLastEditedLink(DataAdapter.Instance.DataCache.Objects(Workspace).Folder_Selected);
        //}

        public void Fav_CreateLastEditedLink()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                //if (_canfav_createlasteditedlink(selectedobject))
                //{
                //    list<idocumentorfolder> selectedobjects = new list<idocumentorfolder>();
                //    selectedobjects.add(selectedobject);
                //    _createfavorlink(csenuminternalobjecttype.favorit, csenuminternalobjecttype.linkobject, csenumprofileworkspace.workspace_favoriten, constants.structlevel_09_doklog, selectedobjects, true);
                //}
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        public bool CanFav_CreateLastEditedLink { get { return _CanFav_CreateLastEditedLink(); } }

        private bool _CanFav_CreateLastEditedLink()
        {
            return _CanFav_CreateLastEditedLink(DataAdapter.Instance.DataCache.Objects(Workspace).Folder_Selected);
        }

        public bool _CanFav_CreateLastEditedLink(IDocumentOrFolder selectedobject)
        {
            //List<IDocumentOrFolder> selectedobjects = new List<IDocumentOrFolder>();
            //selectedobjects.Add(selectedobject);
            //return _CanCreateFoldersInternal(CSEnumInternalObjectType.LinkObject, CSEnumProfileWorkspace.workspace_favoriten, Constants.STRUCTLEVEL_09_DOKLOG, selectedobjects);
            return false;
        }

        // ------------------------------------------------------------------

        #endregion Fav_CreateLastEditedLink (cmd_Fav_CreateLastEditedLink)        

        // ==================================================================

        #region Fav_CreateObjectLink (cmd_Fav_CreateObjectLink)

        public DelegateCommand cmd_Fav_CreateObjectLink { get { return new DelegateCommand(Fav_CreateObjectLink, _CanFav_CreateObjectLink); } }

        public void Fav_CreateObjectLink()
        {
            Fav_CreateObjectLink(DataAdapter.Instance.DataCache.Objects(Workspace).Folder_Selected);
        }

        public void Fav_CreateObjectLink(IDocumentOrFolder selectedobject)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (_CanFav_CreateObjectLink(selectedobject))
                {
                    List<IDocumentOrFolder> selectedobjects = new List<IDocumentOrFolder>();
                    selectedobjects.Add(selectedobject);
                    _CreateFavOrLink(CSEnumInternalObjectType.Favorit, CSEnumInternalObjectType.LinkObject, CSEnumProfileWorkspace.workspace_favoriten, Constants.STRUCTLEVEL_09_DOKLOG, selectedobjects, true);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        public bool CanFav_CreateObjectLink { get { return _CanFav_CreateObjectLink(); } }

        private bool _CanFav_CreateObjectLink()
        {
            return _CanFav_CreateObjectLink(DataAdapter.Instance.DataCache.Objects(Workspace).Folder_Selected);
        }

        public bool _CanFav_CreateObjectLink(IDocumentOrFolder selectedobject)
        {
            List<IDocumentOrFolder> selectedobjects = new List<IDocumentOrFolder>();
            selectedobjects.Add(selectedobject);
            return _CanCreateFoldersInternal(CSEnumInternalObjectType.LinkObject, CSEnumProfileWorkspace.workspace_favoriten, Constants.STRUCTLEVEL_09_DOKLOG, selectedobjects);
        }

        // ------------------------------------------------------------------

        #endregion Fav_CreateObjectLink (cmd_Fav_CreateObjectLink)        

        // ==================================================================

        #region Fav_CreateQueryLink (cmd_Fav_CreateQueryLink)

        public DelegateCommand cmd_Fav_CreateQueryLink { get { return new DelegateCommand(Fav_CreateQueryLink, _CanFav_CreateQueryLink); } }

        public void Fav_CreateQueryLink()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (_CanFav_CreateQueryLink())
                {
                    _CreateFavOrLink(CSEnumInternalObjectType.Favorit, CSEnumInternalObjectType.LinkQuery, CSEnumProfileWorkspace.workspace_favoriten, Constants.STRUCTLEVEL_09_DOKLOG, null, true);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        public bool CanFav_CreateQueryLink { get { return _CanFav_CreateQueryLink(); } }

        public bool _CanFav_CreateQueryLink()
        {
            bool ret = _CanCreateFoldersInternal(CSEnumInternalObjectType.LinkQuery, CSEnumProfileWorkspace.workspace_favoriten, Constants.STRUCTLEVEL_09_DOKLOG, null);
            if (ret) ret = DataAdapter.Instance.DataCache.Objects(this.Workspace).Query_LastQueryValues.Count > 0;
            return ret;
        }

        // ------------------------------------------------------------------

        #endregion Fav_CreateQueryLink (cmd_Fav_CreateQueryLink) 

        // ==================================================================

        #region Fav_CreateURLLink (cmd_Fav_CreateURLLink)

        public DelegateCommand cmd_Fav_CreateURLLink { get { return new DelegateCommand(Fav_CreateURLLink, _CanFav_CreateURLLink); } }

        public void Fav_CreateURLLink()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (_CanFav_CreateURLLink())
                {
                    _CreateFavOrLink(CSEnumInternalObjectType.Favorit, CSEnumInternalObjectType.LinkURL, CSEnumProfileWorkspace.workspace_favoriten, Constants.STRUCTLEVEL_09_DOKLOG, null, true);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        public bool CanFav_CreateURLLink { get { return _CanFav_CreateURLLink(); } }

        public bool _CanFav_CreateURLLink()
        {
            IDocumentOrFolder selObj = DataAdapter.Instance.DataCache.Objects(Workspace).Object_Selected;
            return _CanCreateFoldersInternal(CSEnumInternalObjectType.LinkURL, CSEnumProfileWorkspace.workspace_favoriten, Constants.STRUCTLEVEL_09_DOKLOG, null)
                && (selObj == null || selObj.objectId.Length == 0);
        }

        // ------------------------------------------------------------------

        #endregion Fav_CreateURLLink (cmd_Fav_CreateURLLink)                    


        // ==================================================================

        #region Favs_Show (cmd_Favs_Show)

        public DelegateCommand cmd_Favs_Show { get { return new DelegateCommand(Favs_Show, _CanFavs_Show); } }

        public void Favs_Show()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                // CSProfileComponent profilenode = DataAdapter.Instance.DataCache.Profile.Profile_GetComponentFromLayout(CSEnumProfileComponentType.BASECONTAINER, "Favoriten");
                CSProfileComponent profilenode = DataAdapter.Instance.DataCache.Profile.Profile_GetComponentFromLayout(CSEnumProfileComponentType.FAVORITEN, "Favoriten");
                if (profilenode != null)
                {
                    if (!ViewManager.IsComponentVisible(profilenode))
                    {
                        if (profilenode.keepalive)
                        {
                            ViewManager.ShowComponent(profilenode);
                        }
                        else
                        {
                            ViewManager.LoadComponent(profilenode);
                        }
                    }
                    else
                    {
                        if (profilenode.keepalive)
                            ViewManager.HideComponent(profilenode);
                        else
                            ViewManager.UnloadComponent(profilenode);
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public bool CanFavs_Show { get { return _CanFavs_Show(); } }

        private bool _CanFavs_Show()
        {
            return DataAdapter.Instance.DataCache.ApplicationFullyInit;
        }

        #endregion Favs_Show (cmd_Favs_Show)

        // ==================================================================

        #region Link_CreateFSLink (cmd_Link_CreateFSLink)

        public DelegateCommand cmd_Link_CreateFSLink { get { return new DelegateCommand(Link_CreateFSLink, _CanLink_CreateFSLink); } }

        public void Link_CreateFSLink() { Link_CreateFSLink(null); }

        public void Link_CreateFSLink(string filename)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (_CanLink_CreateFSLink())
                {
                    //_CreateFavOrLink(CSEnumInternalObjectType.Link, CSEnumInternalObjectType.LinkFS, CSEnumProfileWorkspace.workspace_favoriten, Constants.STRUCTLEVEL_09_DOKLOG, true);

                    if (filename == null || filename.Length == 0)
                    {
                        DataAdapter.Instance.DataCache.Info.RemoveInformation(CSEnumInformationId.Undefined);
                        CSInformation info = new CSInformation();
                        info.informationid = CSEnumInformationId.Undefined;
                        info.informationvalue = CSEnumCommands.cmd_Link_CreateFSLink.ToString();
                        CSInformation[] infolist = new CSInformation[1];
                        infolist[0] = info;
                        DataAdapter.Instance.DataCache.Info.AddOrReplaceInformation(infolist);

                        this.LC_FileSelect("");
                    }
                    else
                    {
                        // wenn antwort erhalten dann
                        CSInformation info = new CSInformation();
                        info.informationid = CSEnumInformationId.Undefined;
                        info.informationvalue = filename;
                        CSInformation[] infolist = new CSInformation[1];
                        infolist[0] = info;
                        DataAdapter.Instance.DataCache.Info.AddOrReplaceInformation(infolist);

                        _CreateFavOrLink(CSEnumInternalObjectType.Link, CSEnumInternalObjectType.LinkFS, CSEnumProfileWorkspace.workspace_favoriten, Constants.STRUCTLEVEL_09_DOKLOG, true);
                    }

                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        public bool CanLink_CreateFSLink { get { return _CanLink_CreateFSLink(); } }

        public bool _CanLink_CreateFSLink()
        {
            return CanCreateObjectLink() && _CanCreateFoldersInternal(CSEnumInternalObjectType.LinkFS, CSEnumProfileWorkspace.workspace_favoriten, Constants.STRUCTLEVEL_09_DOKLOG, null);
        }

        // ------------------------------------------------------------------

        #endregion Link_CreateFSLink (cmd_Link_CreateFSLink)        

        // ==================================================================

        #region Link_CreateObjectLink (cmd_Link_CreateObjectLink)

        public DelegateCommand cmd_Link_CreateObjectLink { get { return new DelegateCommand(Link_CreateObjectLink, _CanLink_CreateObjectLink); } }

        public void Link_CreateObjectLink()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (_CanLink_CreateObjectLink())
                {
                    IDocumentOrFolder selObj = DataAdapter.Instance.DataCache.Objects(this.Workspace).Folder_Selected;

                    string linksource = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.ObjectLinkSource];
                    if (linksource != null && linksource.Length > 0)
                    {
                        // es wurde bereits im ersten schritt eine verknüpfung begonnen und im infocache abgelegt
                        // daher kann nun fortgefahren und die relation zwischen beiden angelegt werden
                        string repidsource = linksource.Substring(0, linksource.LastIndexOf("_"));
                        string objectidsource = linksource.Substring(linksource.LastIndexOf("_") + 1);
                        //showOKDialog("Verknüpfung wird erzeugt zwischen: " + repidsource + "." + objectidsource + " und " + selObj.RepositoryId + "." + selObj.objectId);

                        List<IDocumentOrFolder> sourceobjects = DataAdapter.Instance.DataCache.Object_FindObjectWithID(objectidsource);
                        if (sourceobjects == null || sourceobjects.Count == 0)
                        {
                            if(!repidsource.Equals(selObj.RepositoryId))
                            {
                                DataAdapter.Instance.DataProvider.GetObject(repidsource, objectidsource, CSEnumProfileWorkspace.workspace_links, DataAdapter.Instance.DataCache.Rights.UserPrincipal, new CallbackAction(Link_CreateObjectLink_GetSourceObjectDone, repidsource, objectidsource));
                            }
                            else
                            {
                                DataAdapter.Instance.DataProvider.GetObject(repidsource, objectidsource, this.Workspace, DataAdapter.Instance.DataCache.Rights.UserPrincipal, new CallbackAction(Link_CreateObjectLink_GetSourceObjectDone, repidsource, objectidsource));
                            }                            
                        }
                        else
                        {
                            Link_CreateObjectLink_GetSourceObjectDone(repidsource, objectidsource);
                        }
                    }
                    else
                    {
                        //showOKDialog("Anderes Objekt markieren und nochmal CreateObjectLink rufen");
                        //showOKDialog("Bitte wählen Sie das zu verlinkende Objekt im Treeview und wählen erneut ‚Link auf ECM Objekt erstellen‘");
                        if (showOKCancelDialog("Bitte wählen Sie das zu verlinkende Objekt im Treeview und wählen erneut ‚Link auf ECM Objekt erstellen‘"))
                        {
                            string link = selObj.RepositoryId + "_" + selObj.objectId;
                            CSInformation[] infolist = new CSInformation[1];
                            CSInformation info = new CSInformation();
                            info.informationid = CSEnumInformationId.ObjectLinkSource;
                            info.informationvalue = link;
                            infolist[0] = info;
                            DataAdapter.Instance.DataCache.Info.AddOrReplaceInformation(infolist);
                        }
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        private void Link_CreateObjectLink_GetSourceObjectDone(string repidsource, string objectidsource)
        {
            List<IDocumentOrFolder> sourceobjects = DataAdapter.Instance.DataCache.Object_FindObjectWithID(objectidsource);
            IDocumentOrFolder selObj = DataAdapter.Instance.DataCache.Objects(this.Workspace).Folder_Selected;
            if (sourceobjects != null && sourceobjects.Count > 0)
            {
                cmisTypeContainer reltype = DataAdapter.Instance.DataCache.Repository_FindTypeContainerForBaseId(Constants.CMISTYPE_RELATIONSHIP);
                if (reltype != null)
                {
                    // TS 11.08.16
                    if (showOKCancelDialog("Es wird eine Verknüpfung zwischen zwei ECM Objekten erzeugt"))
                    {
                        DataAdapter.Instance.DataProvider.CreateRelationship(reltype, sourceobjects[0], selObj, CSEnumInternalObjectType.LinkObject.ToString(), "Verknüpfung zwischen zwei Objekten",
                        DataAdapter.Instance.DataCache.Rights.UserPrincipal, new CallbackAction(Link_CreateObjectLink_CreateRelationDone, sourceobjects[0], selObj));
                    }
                    // TS 17.08.16
                    else
                        DataAdapter.Instance.DataCache.Info.RemoveInformation(CSEnumInformationId.ObjectLinkSource);
                }

                // Clear Favorite-Cache
                DataAdapter.Instance.Processing(CSEnumProfileWorkspace.workspace_favoriten).ClearCache();
            }
        }

        private void Link_CreateObjectLink_CreateRelationDone(object sourceobject, object targetobject)
        {
            DataAdapter.Instance.DataCache.Info.RemoveInformation(CSEnumInformationId.ObjectLinkSource);

            // beide Objekte neu einlesen damit die Verknüpfung sichtbar wird
            IDocumentOrFolder source = (IDocumentOrFolder)sourceobject;
            IDocumentOrFolder target = (IDocumentOrFolder)targetobject;
            DataAdapter.Instance.DataProvider.GetObject(source.RepositoryId, source.objectId, source.Workspace, DataAdapter.Instance.DataCache.Rights.UserPrincipal);
            DataAdapter.Instance.DataProvider.GetObject(target.RepositoryId, target.objectId, target.Workspace, DataAdapter.Instance.DataCache.Rights.UserPrincipal);
        }

        // ------------------------------------------------------------------
        public bool CanLink_CreateObjectLink { get { return _CanLink_CreateObjectLink(); } }

        public bool _CanLink_CreateObjectLink()
        {
            //return _CanCreateFoldersInternal(CSEnumInternalObjectType.LinkURL, CSEnumProfileWorkspace.workspace_favoriten, Constants.STRUCTLEVEL_09_DOKLOG, null);
            return CanCreateObjectLink();
        }

        // ------------------------------------------------------------------

        #endregion Link_CreateObjectLink (cmd_Link_CreateObjectLink)        

        // ==================================================================

        #region Link_CreateAdressLink (cmd_Link_CreateAdressLink)

        public DelegateCommand cmd_Link_CreateAdressLink { get { return new DelegateCommand(Link_CreateAdressLink, _CanLink_CreateAdressLink); } }

        /// <summary>
        /// Create an Adress<->ECMObject-Link
        /// This one can be triggerd from the silverlight-client or from the adress-client
        /// </summary>
        public void Link_CreateAdressLink()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (_CanLink_CreateAdressLink())
                {
                    IDocumentOrFolder selObj = DataAdapter.Instance.DataCache.Objects(this.Workspace).Folder_Selected;

                    // Case 1: A link-source was already set via the adress-tool and pushed via RC
                    string linksource = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.ObjectAdressLinkSource];
                    if (linksource != null && linksource.Length > 0)
                    {
                        // Extract the id's (original id is like '2C_200_1_O12345')
                        linksource = linksource.Replace("2C_", "");
                        string repidsource = linksource.Substring(0, linksource.LastIndexOf("_"));
                        string objectidsource = linksource.Substring(linksource.LastIndexOf("_") + 1);                        
                        Link_CreateAdressLink_GetSourceObjectDone(repidsource, objectidsource);                        
                    }
                    else
                    {
                        // Case 2: The Action is triggerd from the silverlight-client and a target needs to be set in the adress-tool
                        if (showOKCancelDialog(LocalizationMapper.Instance["msg_dlg_text_createadresslink"]))
                        {                        
                            // Create the 2C-Display-ID
                            string link = "2C_" + selObj.RepositoryId + "_" + selObj.objectId;
                            List<string> dummyList = new List<string>();
                            dummyList.Add(link);

                            // Create the RoutingInfo
                            CSRCPushPullUIRouting routinginfo = new CSRCPushPullUIRouting();
                            CSEnumRCPushUIListeners?[] receiver = new CSEnumRCPushUIListeners?[1];
                            receiver[0] = CSEnumRCPushUIListeners .HTML_CLIENT_ADRESS;
                            routinginfo.sendto_componentorvirtualids = receiver;

                            // Send a PushUI
                            DataAdapter.Instance.DataProvider.RCPush_UI(routinginfo, CSEnumRCPushCommands.newlinktoadress, dummyList, null, DataAdapter.Instance.DataCache.Rights.UserPrincipal);
                        }
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        /// <summary>
        /// This one is called after an adress-client triggered a creation of a link to an ecm-object and the silverlight-client selected an object as link-object.
        /// Here we create the relation itself
        /// </summary>
        /// <param name="repidsource"></param>
        /// <param name="objectidsource"></param>
        private void Link_CreateAdressLink_GetSourceObjectDone(string repidsource, string objectidsource)
        {
            IDocumentOrFolder selObj = DataAdapter.Instance.DataCache.Objects(this.Workspace).Folder_Selected;
            if (repidsource != null && objectidsource != null)
            {
                cmisTypeContainer reltype = DataAdapter.Instance.DataCache.Repository_FindTypeContainerForBaseId(Constants.CMISTYPE_RELATIONSHIP);
                if (reltype != null)
                {
                    if (showOKCancelDialog(LocalizationMapper.Instance["msg_dlg_text_adresslinkcreation"]))
                    {
                        DataAdapter.Instance.DataProvider.CreateRelationshipFromIDs(reltype, objectidsource, repidsource, selObj.objectId, selObj.RepositoryId, CSEnumProfileWorkspace.workspace_links, CSEnumInternalObjectType.AdressObject.ToString(), "Verknüpfung zwischen einem ECM Objekt und einem Adressdatensatz",
                        DataAdapter.Instance.DataCache.Rights.UserPrincipal, new CallbackAction(Link_CreateAdressLink_CreateRelationDone, selObj));
                    }
                    else
                        DataAdapter.Instance.DataCache.Info.RemoveInformation(CSEnumInformationId.ObjectAdressLinkSource);
                }
            }
        }

        /// <summary>
        /// This one is called after a relation is created triggered by the adress-client. 
        /// Here we notify the adress-client about the changes and update the object in our local cache
        /// </summary>
        /// <param name="targetObject"></param>
        private void Link_CreateAdressLink_CreateRelationDone(object targetObject)
        {
            IDocumentOrFolder target = (IDocumentOrFolder)targetObject;       
            string sourceObjectID = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.ObjectAdressLinkSource];
            DataAdapter.Instance.DataCache.Info.RemoveInformation(CSEnumInformationId.ObjectAdressLinkSource);

            // Update the Source-Object to make the relation visible
            DataAdapter.Instance.DataProvider.GetObject(target.RepositoryId, target.objectId, target.Workspace, DataAdapter.Instance.DataCache.Rights.UserPrincipal);

            // Pack the source Display-ID
            List<string> dummyList = new List<string>();
            dummyList.Add(sourceObjectID);

            // Create the RoutingInfo
            CSRCPushPullUIRouting routinginfo = new CSRCPushPullUIRouting();
            CSEnumRCPushUIListeners?[] receiver = new CSEnumRCPushUIListeners?[1];
            receiver[0] = CSEnumRCPushUIListeners.HTML_CLIENT_ADRESS;
            routinginfo.sendto_componentorvirtualids = receiver;

            // Send a PushUI
            DataAdapter.Instance.DataProvider.RCPush_UI(routinginfo, CSEnumRCPushCommands.updobject, dummyList, null, DataAdapter.Instance.DataCache.Rights.UserPrincipal);
        }

        // ------------------------------------------------------------------
        public bool CanLink_CreateAdressLink { get { return _CanLink_CreateAdressLink(); } }

        public bool _CanLink_CreateAdressLink()
        {
            return ViewManager.IsHTMLAdressOpen() && CanCreateObjectLink();
        }

        // ------------------------------------------------------------------

        #endregion Link_CreateAdressLink (cmd_Link_CreateAdressLink)        

        // ==================================================================

        #region Link_CreateURLLink (cmd_Link_CreateURLLink)

        public DelegateCommand cmd_Link_CreateURLLink { get { return new DelegateCommand(Link_CreateURLLink, _CanLink_CreateURLLink); } }

        public void Link_CreateURLLink()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (_CanLink_CreateURLLink())
                {
                    _CreateFavOrLink(CSEnumInternalObjectType.Link, CSEnumInternalObjectType.LinkURL, CSEnumProfileWorkspace.workspace_favoriten, Constants.STRUCTLEVEL_09_DOKLOG, true);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        public bool CanLink_CreateURLLink { get { return _CanLink_CreateURLLink(); } }

        public bool _CanLink_CreateURLLink()
        {
            return CanCreateObjectLink() && _CanCreateFoldersInternal(CSEnumInternalObjectType.LinkURL, CSEnumProfileWorkspace.workspace_favoriten, Constants.STRUCTLEVEL_09_DOKLOG, null);
        }

        // ------------------------------------------------------------------

        #endregion Link_CreateURLLink (cmd_Link_CreateURLLink)        

        // ==================================================================

        #region Link_SendLink (cmd_Link_SendLink)

        public DelegateCommand cmd_Link_SendLink { get { return new DelegateCommand(Link_SendLink, _CanLink_SendLink); } }

        public void Link_SendLink()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (_CanLink_SendLink())
                {
                    string hyperlink = CreateObjectLink(true);
                    LC_MailCreate(null, null, null, null, null, hyperlink, null, null, false, null, false);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        public bool CanLink_SendLink { get { return _CanLink_SendLink(); } }

        public bool _CanLink_SendLink() { return CanCreateObjectLink(); }

        // ------------------------------------------------------------------

        #endregion Link_SendLink (cmd_Link_SendLink)      

        // ==================================================================

        #region Create Object Link (cmd_CreateObjectLink)

        public DelegateCommand cmd_CreateObjectLink { get { return new DelegateCommand(CreateObjectLink, _CanCreateObjectLink); } }

        private bool _CanCreateObjectLink()
        {
            return CanCreateObjectLink();
        }

        public bool CanCreateObjectLink()
        {
            return DataAdapter.Instance.DataCache.Objects(this.Workspace).Object_Selected != null
                && DataAdapter.Instance.DataCache.Objects(this.Workspace).Object_Selected.objectId.Length > 0
                && DataAdapter.Instance.DataCache.Objects(this.Workspace).Object_Selected.isFolder;
        }

        /// <summary>
        /// Creates a command-based hyperlink for webservice-calls
        /// </summary>
        public void CreateObjectLink()
        {
            try
            {
                if (CanCreateObjectLink())
                {
                    // TS 10.08.16
                    // TS 05.07.16 das geht nur für ordner glaube ich
                    //IDocumentOrFolder selObj = DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_default).Object_Selected;
                    IDocumentOrFolder selObj = DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_default).Folder_Selected;
                    // TS 10.08.16
                    //string textToSet = DataAdapter.Instance.DataProvider.GetProviderAttribs(true, true, true, true, false);

                    //// Building the final link
                    //// TODO: Add possible further params here (like documents, pages, etc.)
                    //string paramLink = "2C_" + selObj.RepositoryId + "_" + selObj.objectId;
                    //textToSet += "?display="+ HttpUtility.UrlEncode("b64~" + Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(paramLink)));

                    string hyperlink = CreateObjectLink(true);
                    // TS 11.08.16 direkt in die zwischenablage
                    // Create Dialog-Window aaand show it
                    //View.Dialogs.dlgCreateObjectLink dlgCreateObjLink = new View.Dialogs.dlgCreateObjectLink(this.Workspace, selObj, textToSet);
                    //dlgCreateObjLink.Show();
                    System.Windows.Clipboard.SetText(hyperlink);
                }
            }
            catch (Exception ex) { if (ex.Message.Length > 0) { Log.Log.Debug(ex.Message); } }
        }

        private string CreateObjectLink(bool encodeb64)
        {
            string textToSet = "";
            try
            {
                if (CanCreateObjectLink())
                {
                    // TS 05.07.16 das geht nur für ordner glaube ich
                    //IDocumentOrFolder selObj = DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_default).Object_Selected;
                    IDocumentOrFolder selObj = DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_default).Folder_Selected;
                    textToSet = DataAdapter.Instance.DataProvider.GetProviderAttribs(true, true, true, true, false);

                    // Building the final link
                    // TODO: Add possible further params here (like documents, pages, etc.)
                    string paramLink = "2C_" + selObj.RepositoryId + "_" + selObj.objectId;
                    if (encodeb64)
                    {
                        textToSet += "?display=" + HttpUtility.UrlEncode("b64~" + Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(paramLink)));
                    }
                    else
                        textToSet += "?display=" + paramLink;
                }
            }
            catch (Exception ex) { if (ex.Message.Length > 0) { Log.Log.Debug(ex.Message); } }
            return textToSet;
        }

        #endregion Create Object Link (cmd_CreateObjectLink)


        // ==================================================================

        #region CreateSustainableFileLinks (cmd_CreateSustainableFileLinks)

        public DelegateCommand cmd_CreateSustainableFileLinks { get { return new DelegateCommand(CreateSustainableFileLinks, _CanCreateSustainableFileLinks); } }

        private void CreateSustainableFileLinks()
        {
            CreateSustainableFileLinks(DataAdapter.Instance.DataCache.Objects(Workspace).Objects_Selected);
        }

        public void CreateSustainableFileLinks(List<IDocumentOrFolder> objectlist)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (_CanCreateSustainableFileLinks(objectlist))
                {
                    List<IDocumentOrFolder> createobjects = new List<IDocumentOrFolder>();
                    List<string> createdids = new List<string>();
                    List<cmisObjectType> createcmisobjects = new List<cmisObjectType>();
                    foreach (IDocumentOrFolder obj in objectlist)
                    {
                        switch (obj.structLevel)
                        {
                            case Constants.STRUCTLEVEL_09_DOKLOG:
                                createobjects.AddRange(obj.GetNewestChildDokVersions(false));
                                if (createobjects.Count > 0)
                                {
                                    createdids.Add(createobjects[0].objectId);
                                    createcmisobjects.Add(createobjects[0].CMISObject);
                                }
                                break;

                            case Constants.STRUCTLEVEL_10_DOKUMENT:
                                createdids.Add(obj.objectId);
                                createcmisobjects.Add(obj.CMISObject);
                                break;

                            case Constants.STRUCTLEVEL_11_FULLTEXT:
                                createdids.Add(obj.objectId);
                                createcmisobjects.Add(obj.CMISObject);
                                break;
                        }
                    }
                    if (createcmisobjects.Count > 0)
                    {
                        CallbackAction callback = new CallbackAction(CreateSustainableFileLinks_Done, createdids);
                        DataAdapter.Instance.DataProvider.CreateSustainableFileLinks(createcmisobjects, DataAdapter.Instance.DataCache.Rights.UserPrincipal, callback);
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        public bool CanCreateSustainableFileLinks { get { return _CanCreateSustainableFileLinks(DataAdapter.Instance.DataCache.Objects(Workspace).Objects_Selected); } }

        private bool _CanCreateSustainableFileLinks()
        {
            return _CanCreateSustainableFileLinks(DataAdapter.Instance.DataCache.Objects(Workspace).Objects_Selected);
        }

        private bool _CanCreateSustainableFileLinks(List<IDocumentOrFolder> objectlist)
        {
            bool cancreate = DataAdapter.Instance.DataCache.ApplicationFullyInit && objectlist != null && objectlist.Count > 0;
            if (cancreate)
            {
                foreach (IDocumentOrFolder obj in objectlist)
                {
                    cancreate = !obj.isNotCreated && obj.objectId.Length > 0;
                    cancreate = cancreate && ((obj.structLevel == Constants.STRUCTLEVEL_09_DOKLOG && obj.hasChildDocuments) || obj.structLevel > Constants.STRUCTLEVEL_09_DOKLOG);
                    if (!cancreate)
                        break;
                }
            }
            return cancreate;
        }

        ///// <summary>
        ///// info an alle
        ///// </summary>
        private void CreateSustainableFileLinks_Done(object createobjectslist)
        {
            // List<IDocumentOrFolder> createobjects
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            if (DataAdapter.Instance.DataCache.ResponseStatus.success)
            {
                if (createobjectslist != null && createobjectslist.GetType().Name.StartsWith("List"))
                {
                    List<string> createobjects = (List<string>)createobjectslist;
                    string content = "";
                    foreach (string objectid in createobjects)
                    {
                        IDocumentOrFolder obj = DataAdapter.Instance.DataCache.Objects(this.Workspace).GetObjectById(objectid);
                        if (obj != null && obj.objectId.Equals(objectid))
                        {
                            string url = obj[CSEnumCmisProperties.sustainableDownloadName];
                            string date = obj[CSEnumCmisProperties.sustainableDownloadExpireDate];

                            content = content + "Verknüpfung gültig bis: " + date + "   " + url + "\r\n";
                            //content = content + "Verknüpfung gültig bis: " + date + "   " + url + "\n";
                            //content = content + "Hallo nächste Zeile";
                            //content = content + "Verknüpfung gültig bis: " + date + "   " + url + "\r u13\u10";
                        }
                    }
                    LC_MailCreate(null, null, null, null, null, content, null, null, false, null);
                }
            }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        #endregion CreateSustainableFileLinks (cmd_CreateSustainableFileLinks)

        // ==================================================================

        private void _CreateFavOrLink(CSEnumInternalObjectType objecttype, CSEnumInternalObjectType objectsubtype, CSEnumProfileWorkspace workspace, int structlevel, bool showdialog)
        {
            _CreateFavOrLink(objecttype, objectsubtype, workspace, structlevel, DataAdapter.Instance.DataCache.Objects(this.Workspace).Objects_Selected, showdialog);
        }

        private void _CreateFavOrLink(CSEnumInternalObjectType objecttype, CSEnumInternalObjectType objectsubtype, CSEnumProfileWorkspace workspace, int structlevel, List<IDocumentOrFolder> mainselectedobjects, bool showdialog)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (_CanCreateFoldersInternal(objecttype, workspace, structlevel, mainselectedobjects))
                {
                    cmisTypeContainer typedef = DataAdapter.Instance.DataCache.Repository(workspace).GetTypeContainerForStructLevel(structlevel);
                    cmisObjectType parentobject = DataAdapter.Instance.DataCache.Objects(workspace).Root.CMISObject;
                    CServer.cmisTypeContainer typedefrelation = null;
                    List<cmisObjectType> relationobjects = new List<cmisObjectType>();

                    bool autosave = !showdialog;

                    typedefrelation = DataAdapter.Instance.DataCache.Repository(Workspace).GetTypeContainerForBaseId("cmisrelationship");
                    if (typedefrelation == null) typedefrelation = DataAdapter.Instance.DataCache.Repository_FindTypeContainerForBaseId("cmisrelationship");

                    if (DataAdapter.Instance.DataCache.Objects(workspace).Object_Selected.objectId.Length == 0)
                    {
                        DataAdapter.Instance.DataCache.Objects(workspace).SetSelectedObject(DataAdapter.Instance.DataCache.Objects(workspace).Root);
                        // die dummy objekte zur suche wegnehmen da ja nun ein echter ordner angelegt wird
                        DataAdapter.Instance.DataCache.Objects(workspace).RemoveEmptyQueryObjects(true);
                    }

                    // TS 23.04.15
                    if (autosave) typedef = AddDefaultValuesToTypeContainer(typedef, false);

                    bool addrelation = false;
                    if (!addrelation) addrelation = objecttype == CSEnumInternalObjectType.Favorit && objectsubtype == CSEnumInternalObjectType.LinkObject;
                    if (!addrelation) addrelation = objecttype == CSEnumInternalObjectType.Link && objectsubtype == CSEnumInternalObjectType.LinkFS;
                    if (!addrelation) addrelation = objecttype == CSEnumInternalObjectType.Link && objectsubtype == CSEnumInternalObjectType.LinkURL;
                    if (addrelation) addrelation = mainselectedobjects != null && mainselectedobjects.Count > 0;

                    if (addrelation)
                    {
                        relationobjects.Add(mainselectedobjects[0].CMISObject);
                    }

                    DataAdapter.Instance.DataProvider.CreateFoldersInternal(typedef, objectsubtype, parentobject, typedefrelation, relationobjects, autosave, DataAdapter.Instance.DataCache.Rights.UserPrincipal, new CallbackAction(_CreateFavOrLink_Done, objecttype, objectsubtype, workspace, mainselectedobjects));
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // private void _CreateFavOrLink_Done(CSEnumInternalObjectType objectsubtype, CSEnumProfileWorkspace workspace, List<IDocumentOrFolder> mainselectedobjects)
        private void _CreateFavOrLink_Done(object oobjecttype, object oobjectsubtype, object oworkspace, object omainselectedobjects)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                // CSEnumInternalObjectType objectsubtype, CSEnumProfileWorkspace workspace, List<IDocumentOrFolder> mainselectedobjects
                CSEnumInternalObjectType objecttype = (CSEnumInternalObjectType)oobjecttype;
                CSEnumInternalObjectType objectsubtype = (CSEnumInternalObjectType)oobjectsubtype;
                CSEnumProfileWorkspace workspace = (CSEnumProfileWorkspace)oworkspace;
                List<IDocumentOrFolder> mainselectedobjects = (List<IDocumentOrFolder>)omainselectedobjects;

                bool valid = false;
                if (DataAdapter.Instance.DataCache.ResponseStatus.success)
                {
                    // die erstellten objekte zusammensuchen: 1. das letzte in der liste ist das favoriten-objekt, das vorletzte ist das link-objekt, beide sollten über eine relation verknüpft sein
                    if (DataAdapter.Instance.DataCache.Objects(workspace).ObjectList.Count > 1)
                    {
                        IDocumentOrFolder favobject = null;
                        IDocumentOrFolder linkobject = null;

                        if (objecttype == CSEnumInternalObjectType.Favorit)
                        {
                            favobject = DataAdapter.Instance.DataCache.Objects(workspace).ObjectList[DataAdapter.Instance.DataCache.Objects(workspace).ObjectList.Count - 1];
                            if (!objectsubtype.ToString().Equals(CSEnumInternalObjectType.LinkObject.ToString()))
                            {
                                linkobject = DataAdapter.Instance.DataCache.Objects(workspace).ObjectList[DataAdapter.Instance.DataCache.Objects(workspace).ObjectList.Count - 2];
                            }
                            if (favobject != null && favobject.Relationships.Count > 0)
                            {
                                DocOrFolderRelationship favrelation = favobject.Relationships[0];
                                if (favrelation.source_objectid.Equals(favobject.objectId) && (linkobject == null || favrelation.target_objectid.Equals(linkobject.objectId)))
                                {
                                    valid = true;
                                    List<string> idlist = new List<string>();
                                    idlist.Add(favobject.objectId);
                                    if (linkobject != null) idlist.Add(linkobject.objectId);
                                    DataAdapter.Instance.Processing(workspace).SetSelectedObject(favobject.objectId);
                                    DataAdapter.Instance.Processing(workspace).SetSelectedObjects(idlist);
                                }
                            }
                        }
                        else if (objecttype == CSEnumInternalObjectType.Link)
                        {
                            linkobject = DataAdapter.Instance.DataCache.Objects(workspace).ObjectList[DataAdapter.Instance.DataCache.Objects(workspace).ObjectList.Count - 1];
                            if (linkobject != null && linkobject.Relationships.Count > 0)
                            {
                                DocOrFolderRelationship linkrelation = linkobject.Relationships[0];
                                if (linkrelation.target_objectid.Equals(linkobject.objectId))
                                {
                                    valid = true;
                                    //List<string> idlist = new List<string>();
                                    //idlist.Add(linkobject.objectId);
                                    DataAdapter.Instance.Processing(workspace).SetSelectedObject(linkobject.objectId);
                                    //DataAdapter.Instance.Processing(workspace).SetSelectedObjects(idlist);
                                }
                            }
                        }

                    }
                }

                if (valid)
                {
                    SL2C_Client.View.Dialogs.dlgCreateFavOrLink child = new View.Dialogs.dlgCreateFavOrLink(workspace, mainselectedobjects);
                    DialogHandler.Show_Dialog(child);
                }
                else
                {
                    DataAdapter.Instance.Processing(workspace).Cancel();
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ==================================================================

        #region MergeFiles (cmd_MergeFiles)

        public DelegateCommand cmd_MergeFiles { get { return new DelegateCommand(MergeFiles, _CanMergeFiles); } }

        public void MergeFiles()
        {
            // TS 23.06.15 überarbeitet um auch aus der ablagemappe ungespeicherte (kopierte) verarbeiten zu können
            // und um die dateiversionen hier und nicht im Silverlight zu ermitteln
            //
            //														deep	paginate	input
            //Export			=> vollständiger rekursiver Export	true	true		all levels
            //
            //Print				=> nur Dateien und Dokumente		false	false		> = 9
            //
            //MergeFiles    	=> wie print + automatisch ablegen	false	false		> = 9
            //
            //{ConvertToPDF]	=> CServer intern, nach scannen		false	false		> = 10
            //
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanMergeFiles)
                {
                    MergeFiles(DataAdapter.Instance.DataCache.Objects(this.Workspace).Objects_Selected);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public void MergeFiles(List<IDocumentOrFolder> objectlist)
        {
            // TS 23.06.15 überarbeitet um auch aus der ablagemappe ungespeicherte (kopierte) verarbeiten zu können
            // und um die dateiversionen hier und nicht im Silverlight zu ermitteln
            //
            //														deep	paginate	input
            //Export			=> vollständiger rekursiver Export	true	true		all levels
            //
            //Print				=> nur Dateien und Dokumente		false	false		> = 9
            //
            //MergeFiles    	=> wie print + automatisch ablegen	false	false		> = 9
            //
            //{ConvertToPDF]	=> CServer intern, nach scannen		false	false		> = 10
            //
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (_CanMergeFiles(objectlist))
                {
                    Export(objectlist, false, false, false, new CallbackAction(MergeFiles_Done));
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        public bool CanMergeFiles { get { return _CanMergeFiles(DataAdapter.Instance.DataCache.Objects(Workspace).Objects_Selected); } }

        private bool _CanMergeFiles()
        {
            return _CanMergeFiles(DataAdapter.Instance.DataCache.Objects(Workspace).Objects_Selected);
        }

        private bool _CanMergeFiles(List<IDocumentOrFolder> objects)
        {
            bool ret = DataAdapter.Instance.DataCache.ApplicationFullyInit && objects.Count > 0;
            if (ret)
            {
                foreach (IDocumentOrFolder obj in objects)
                {
                    if (ret) ret = obj.hasChildDocuments || obj.isDocument;
                    if (!ret)
                        break;
                }
            }
            return ret;
        }

        private void MergeFiles_Done()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (DataAdapter.Instance.DataCache.ResponseStatus.success)
                {
                    string filenamepdf = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.URL_ExportPDF];
                    //LC_Print(filenamepdf);
                    // in ablagemappe packen
                    if (filenamepdf != null && filenamepdf.Length > 0 && filenamepdf.Contains("/"))
                    {
                        cmisContentStreamType content = new cmisContentStreamType();

                        // TS 08.04.14
                        //string shortname = filenamepdf.Substring(filenamepdf.LastIndexOf("/"));
                        //content.filename = shortname;
                        //content.mimeType = shortname;
                        char splitter = ("/".ToCharArray())[0];
                        string[] tokens = filenamepdf.Split(splitter);
                        content.filename = tokens[tokens.Count() - 1];
                        content.mimeType = Statics.Constants.CONTENTSTREAM_MIMETYPE_FLAG_INTERNAL_DOC + tokens[tokens.Count() - 2];
                        
                        List<cmisContentStreamType> contents = new List<cmisContentStreamType>();
                        contents.Add(content);

                        // Always attach to clipboard!
                        IDocumentOrFolder parent = DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_clipboard).Root;
                        cmisTypeContainer typedef = DataAdapter.Instance.DataCache.Repository_FindTypeContainerForBaseId("cmisdocument");
                        DataAdapter.Instance.DataProvider.CreateDocuments(typedef, parent.CMISObject, contents, CSEnumDocumentCopyTypes.COPYTYPE_ORIGINAL, "", false, DataAdapter.Instance.DataCache.Rights.UserPrincipal, new CallbackAction(CreateDocuments_Done));
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        #endregion MergeFiles (cmd_MergeFiles)

        // ==================================================================

        #region Print (cmd_Print)

        public DelegateCommand cmd_Print { get { return new DelegateCommand(Print, _CanPrint); } }

        public void Print()
        {
            // TS 23.06.15 überarbeitet um auch aus der ablagemappe ungespeicherte (kopierte) verarbeiten zu können
            // und um die dateiversionen hier und nicht im Silverlight zu ermitteln
            //
            //														deep	paginate	input
            //Export			=> vollständiger rekursiver Export	true	true		all levels
            //
            //Print				=> nur Dateien und Dokumente		false	false		> = 9
            //
            //MergeFiles    	=> wie print + automatisch ablegen	false	false		> = 9
            //
            //{ConvertToPDF]	=> CServer intern, nach scannen		false	false		> = 10
            //
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanPrint)
                {
                    Print(DataAdapter.Instance.DataCache.Objects(this.Workspace).Objects_Selected);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public void Print(List<IDocumentOrFolder> objectlist)
        {
            // TS 23.06.15 überarbeitet um auch aus der ablagemappe ungespeicherte (kopierte) verarbeiten zu können
            // und um die dateiversionen hier und nicht im Silverlight zu ermitteln
            //
            //														deep	paginate	input
            //Export			=> vollständiger rekursiver Export	true	true		all levels
            //
            //Print				=> nur Dateien und Dokumente		false	false		> = 9
            //
            //MergeFiles    	=> wie print + automatisch ablegen	false	false		> = 9
            //
            //{ConvertToPDF]	=> CServer intern, nach scannen		false	false		> = 10
            //
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (_CanPrint(objectlist))
                {
                    // TS 25.06.15 wenn es nur eine datei ist und die ist bereits pdf dann einfach öffnen
                    if (objectlist.Count == 1 && objectlist[0].DownloadNamePDF != null && objectlist[0].DownloadNamePDF.Length > 0)
                    {
                        // TS 21.12.15
                        //System.Windows.Browser.HtmlWindow newWindow;
                        //newWindow = HtmlPage.Window.Navigate(new Uri(objectlist[0].DownloadNamePDF, UriKind.RelativeOrAbsolute), "_blank", "");
                        DisplayExternal(objectlist[0].DownloadNamePDF);
                    }
                    else
                        Export(objectlist, false, false, false, new CallbackAction(Print_Done));
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        public bool CanPrint { get { return _CanPrint(DataAdapter.Instance.DataCache.Objects(Workspace).Objects_Selected); } }

        private bool _CanPrint()
        {
            return _CanPrint(DataAdapter.Instance.DataCache.Objects(Workspace).Objects_Selected);
        }

        private bool _CanPrint(List<IDocumentOrFolder> objects)
        {
            bool ret = DataAdapter.Instance.DataCache.ApplicationFullyInit && objects.Count > 0;
            if (ret)
            {
                foreach (IDocumentOrFolder obj in objects)
                {
                    if (ret) ret = obj.hasChildDocuments || obj.isDocument;
                    if (!ret)
                        break;
                }
            }
            return ret;
        }

        private void Print_Done()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (DataAdapter.Instance.DataCache.ResponseStatus.success)
                {
                    string filenamepdf = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.URL_ExportPDF];
                    //LC_Print(filenamepdf);
                    // extern anzeigen, dort kann dann gedruckt werden
                    if (filenamepdf != null && filenamepdf.Length > 0)
                    {
                        // TS 21.12.15
                        //System.Windows.Browser.HtmlWindow newWindow;
                        //newWindow = HtmlPage.Window.Navigate(new Uri(filenamepdf, UriKind.RelativeOrAbsolute), "_blank", "");
                        DisplayExternal(filenamepdf);
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        #endregion Print (cmd_Print)

        // ==================================================================

        #region GetAllGroups

        public void GetAllGroups()
        {
            GetAllGroups(null);
        }

        public void GetAllGroups(CallbackAction finalcallback)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanGetAllGroups)
                {
                    DataAdapter.Instance.DataProvider.GetAllGroups(DataAdapter.Instance.DataCache.Rights.UserPrincipal, finalcallback);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        public bool CanGetAllGroups { get { return _CanGetAllGroups(); } }

        public bool _CanGetAllGroups()
        {
            return (DataAdapter.Instance.DataCache.ApplicationFullyInit);
        }

        #endregion GetAllGroups

        // ==================================================================

        #region GetAllUsers

        public void GetAllUsers()
        {
            GetAllUsers(null);
        }

        public void GetAllUsers(CallbackAction finalcallback)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanGetAllUsers)
                {
                    DataAdapter.Instance.DataProvider.GetAllUsers(DataAdapter.Instance.DataCache.Rights.UserPrincipal, finalcallback);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        public bool CanGetAllUsers { get { return _CanGetAllUsers(); } }

        public bool _CanGetAllUsers()
        {
            return (DataAdapter.Instance.DataCache.ApplicationFullyInit);
        }

        #endregion GetAllUsers

        // ==================================================================

        #region GetChildCount

        public void GetChildCount(IDocumentOrFolder objForProc, CallbackAction finalcallback)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanGetChildCount)
                {
                    DataAdapter.Instance.DataProvider.GetChildCount(objForProc.RepositoryId, objForProc.objectId, DataAdapter.Instance.DataCache.Rights.UserPrincipal, finalcallback);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        public bool CanGetChildCount { get { return _CanGetChildCount(); } }

        public bool _CanGetChildCount()
        {
            return (DataAdapter.Instance.DataCache.ApplicationFullyInit);
        }

        #endregion GetChildCount

        // ==================================================================

        #region GetChildren (cmd_GetChildren)

        public DelegateCommand cmd_GetChildren { get { return new DelegateCommand(_GetChildrenViaCommand, _CanGetChildren); } }

        public void GetChildren(string objectid, CallbackAction finalcallback)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (objectid == null) objectid = "";
                if (objectid.Length == 0) objectid = DataAdapter.Instance.DataCache.Objects(Workspace).Folder_Selected.objectId;

                if (_CanGetChildren(objectid))
                {
                    IDocumentOrFolder tmp = DataAdapter.Instance.DataCache.Objects(Workspace).GetObjectById(objectid);
                    GetObjectsUpDown(tmp, true, finalcallback);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public void GetChildren()
        {
            GetChildren("");
        }

        public void GetChildren(string objectid)
        {
            GetChildren(objectid, null);
        }

        private void _GetChildrenViaCommand(object parameter)
        {

            CallbackAction cb = new CallbackAction(SelectFirstChild, (string)parameter);
            GetChildren((string)parameter, cb);
        }

        private void SelectFirstChild(string parentObjectId)
        {
            IDocumentOrFolder parentObj = DataAdapter.Instance.DataCache.Objects(Workspace).GetObjectById(parentObjectId);
            if(parentObj != null && parentObj.objectId.Length > 0 && parentObj.ChildObjects.Count > 0)
            {
                DataAdapter.Instance.DataCache.Info.TreeView_ExpandItemWithID = parentObj.objectId;
                DataAdapter.Instance.InformObservers();
            }
        }

        // ------------------------------------------------------------------
        public bool CanGetChildren { get { return _CanGetChildren(""); } }

        private bool _CanGetChildren(object parameter)
        {
            return _CanGetChildren((string)parameter);
        }

        private bool _CanGetChildren(string objectid)
        {
            if (objectid == null) objectid = "";
            if (objectid.Length == 0) objectid = DataAdapter.Instance.DataCache.Objects(Workspace).Folder_Selected.objectId;
            IDocumentOrFolder dummy = DataAdapter.Instance.DataCache.Objects(Workspace).GetObjectById(objectid);
            return (DataAdapter.Instance.DataCache.ApplicationFullyInit && dummy.canGetChildren && dummy.hasMoreItems);
        }

        #endregion GetChildren (cmd_GetChildren)

        // ==================================================================

        #region GetDocumentPages (cmd_GetDocumentPages)

        public DelegateCommand cmd_GetDocumentPages { get { return new DelegateCommand(_GetDocumentPages, _CanGetDocumentPages); } }

        public void GetDocumentPages(string objectid, int callbacklevel, CallbackAction finalcallback)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                // TS 13.12.17 nach oben weil in beiden callbacklevels benötigt
                if (objectid == null) objectid = "";
                if (objectid.Length == 0)
                {
                    objectid = DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected.objectId;
                }
                IDocumentOrFolder tmp = DataAdapter.Instance.DataCache.Objects(Workspace).GetObjectById(objectid);

                if (callbacklevel == 0)
                {
                    if (_CanGetDocumentPages(objectid))
                    {
                        // TS 13.12.17 nach oben weil in beiden callbacklevels benötigt                     
                        //IDocumentOrFolder tmp = null;
                        //if (objectid == null) objectid = "";
                        //if (objectid.Length == 0)
                        //    tmp = DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected;
                        //else
                        //    tmp = DataAdapter.Instance.DataCache.Objects(Workspace).GetObjectById(objectid);

                        if (tmp != null)
                        {
                            int pagefrom = tmp.DZPageCountConverted;
                            string maxdownloads = DataAdapter.Instance.DataCache.Profile.Option_GetValue(CSEnumOptions.maxdownloads);
                            int pageto = pagefrom + int.Parse(maxdownloads);
                            if (pageto > tmp.DZPageCount)
                                pageto = tmp.DZPageCount;

                            cmisObjectType doc = tmp.CMISObject;
                            IDocumentOrFolder tmpparent = DataAdapter.Instance.DataCache.Objects(Workspace).GetObjectById(tmp.TempOrRealParentId);

                            cmisObjectType parent = tmpparent.CMISObject;
                            List<CServer.cmisObjectType> cmisdocuments = new List<CServer.cmisObjectType>();
                            _CollectDZImagesRecursive(tmpparent, cmisdocuments);

                            // TS 01.02.16 auf nächste gelesene Seite positionieren
                            DataAdapter.Instance.DataProvider.GetDocumentPages(doc, parent, cmisdocuments, pagefrom, pageto, DataAdapter.Instance.DataCache.Rights.UserPrincipal, new CallbackAction(GetDocumentPages, objectid, pagefrom, finalcallback));
                        }
                    }
                    else
                        if (finalcallback != null) finalcallback.Invoke();
                }
                // TS 01.02.16 auf nächste gelesene Seite positionieren
                else if (callbacklevel > 0)
                {
                    if (DataAdapter.Instance.DataCache.ResponseStatus.success)
                    {
                        // TS 13.12.17                        
                        //DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected.SelectedPage = callbacklevel;
                        //DataAdapter.Instance.DataCache.Objects(Workspace).SetSelectedObject(DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected.objectId);
                        tmp.SelectedPage = callbacklevel;
                        DataAdapter.Instance.DataCache.Objects(Workspace).SetSelectedObject(objectid);

                        DataAdapter.Instance.InformObservers();

                        if (finalcallback != null) finalcallback.Invoke();
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        //public void GetDocumentPages(string objectid, int callbacklevel, CallbackAction finalcallback)
        //{
        //    if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
        //    try
        //    {
        //        if (callbacklevel == 0)
        //        {
        //            if (objectid == null) objectid = "";
        //            if (objectid.Length == 0) objectid = DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected.objectId;

        //            // DZPageCount > 0 && DZPageCountConverted
        //            //<option editable="true" id="maxdownloads" value="5" visible="true"/>

        //            int pagefrom = DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected.DZPageCountConverted;
        //            string maxdownloads = DataAdapter.Instance.DataCache.Profile.Option_GetValue(CSEnumOptions.maxdownloads);
        //            int pageto = pagefrom + int.Parse(maxdownloads);
        //            if (pageto > DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected.DZPageCount)
        //                pageto = DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected.DZPageCount;

        //            if (_CanGetDocumentPages(objectid))
        //            {
        //                CServer.cmisObjectType doc = DataAdapter.Instance.DataCache.Objects(Workspace).GetObjectById(objectid).CMISObject;
        //                DataAdapter.Instance.DataProvider.GetDocumentPages(doc, pagefrom, pageto, DataAdapter.Instance.DataCache.Rights.UserPrincipal, new CallbackAction(GetDocumentPages, objectid, 1, finalcallback));
        //            }
        //            else
        //                if (finalcallback != null) finalcallback.Invoke();
        //        }
        //        else if (callbacklevel == 1)
        //        {
        //            if (DataAdapter.Instance.DataCache.ResponseStatus.success)
        //            {
        //                DataAdapter.Instance.DataCache.Objects(Workspace).SetSelectedObject(DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected.objectId);
        //                DataAdapter.Instance.InformObservers();

        //                if (finalcallback != null) finalcallback.Invoke();
        //            }
        //        }
        //    }
        //    catch (Exception e) { Log.Log.Error(e); }
        //    if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        //}
        public void GetDocumentPages() { GetDocumentPages(""); }

        public void GetDocumentPages(string objectid)
        {
            GetDocumentPages(objectid, 0, null);
        }

        private void _GetDocumentPages(object parameter)
        {
            GetDocumentPages((string)parameter);
        }

        // ------------------------------------------------------------------
        public bool CanGetDocumentPages { get { return _CanGetDocumentPages(""); } }

        private bool _CanGetDocumentPages(object parameter)
        {
            return _CanGetDocumentPages((string)parameter);
        }

        private bool _CanGetDocumentPages(string objectid)
        {
            return DataAdapter.Instance.DataCache.ApplicationFullyInit;
        }

        #endregion GetDocumentPages (cmd_GetDocumentPages)

        // ==================================================================

        #region GetInformationList

        public void GetInformationList(List<CSInformation> infoasked, CallbackAction finalcallback)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (_CanGetInformationList(infoasked))
                {
                    // TS 10.11.14 das callback fehlte
                    // DataAdapter.Instance.DataProvider.GetInformationList(infoasked, DataAdapter.Instance.DataCache.Rights.UserPrincipal);
                    DataAdapter.Instance.DataProvider.GetInformationList(infoasked, DataAdapter.Instance.DataCache.Rights.UserPrincipal, finalcallback);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // TS 10.02.16
        public void GetInformationList() { GetInformationList(null); }

        public void GetInformationList(CallbackAction finalcallback)
        {
            // TS 17.09.13
            List<CSInformation> infosasked = new List<CSInformation>();

            CSInformation info = new CSInformation();
            info.informationid = CSEnumInformationId.Unread_PostCount;
            info.informationidSpecified = true;
            infosasked.Add(info);

            info = new CSInformation();
            info.informationid = CSEnumInformationId.Unread_MailCount;
            info.informationidSpecified = true;
            infosasked.Add(info);

            //info = new CSInformation();
            //info.informationid = CSEnumInformationId.Current_TerminCount;
            //info.informationidSpecified = true;
            //infosasked.Add(info);

            info = new CSInformation();
            info.informationid = CSEnumInformationId.Current_AufgabeCount;
            info.informationidSpecified = true;
            infosasked.Add(info);

            info = new CSInformation();
            info.informationid = CSEnumInformationId.Current_StapelCount;
            info.informationidSpecified = true;
            infosasked.Add(info);

            info = new CSInformation();
            info.informationid = CSEnumInformationId.Current_AblageCount;
            info.informationidSpecified = true;
            infosasked.Add(info);

            // nicht abfragen, wird automatisch ermittelt
            //info = new CSInformation();
            //info.informationid = CSEnumInformationId.Overall_ToDoCount;
            //info.informationidSpecified = true;
            //infosasked.Add(info);

            // TS 04.05.16 das timeout mitlesen wenn noch nicht gelesen
            if (DataAdapter.Instance.DataCache.Info[CSEnumInformationId.CommunTimeout] == null || DataAdapter.Instance.DataCache.Info[CSEnumInformationId.CommunTimeout].Length == 0)
            {
                info = new CSInformation();
                info.informationid = CSEnumInformationId.CommunTimeout;
                info.informationidSpecified = true;
                infosasked.Add(info);
            }

            // TS 10.11.14
            //GetInformationList(infosasked, null);
            GetInformationList(infosasked, finalcallback);
        }

        //public void GetInformationList(CallbackAction finalcallback) { GetInformationList(null, finalcallback); }
        // ------------------------------------------------------------------
        public bool CanGetInformationList { get { return _CanGetInformationList(null); } }

        private bool _CanGetInformationList(List<CSInformation> infoasked)
        {
            return DataAdapter.Instance.DataCache.ApplicationFullyInit;
        }

        #endregion GetInformationList

        // ==================================================================

        #region GetObjectsUpDown

        // TS 03.02.16

        private void GetObjectsUpDown_CB(object docorfolder, object bool_forcereadchildren, CallbackAction finalcallback)
        {
            GetObjectsUpDown((IDocumentOrFolder)docorfolder, (bool)bool_forcereadchildren, finalcallback);
        }

        public bool GetObjectsUpDown(IDocumentOrFolder docorfolder, bool forcereadchildren, CallbackAction finalcallback)
        {
            return GetObjectsUpDown(docorfolder, forcereadchildren, false, finalcallback);
        }

        public bool GetObjectsUpDown(object docorfolder, bool forcereadchildren, bool forcegetdzcollection, CallbackAction finalcallback)
        {
            return GetObjectsUpDown((IDocumentOrFolder)docorfolder, forcereadchildren, forcegetdzcollection, finalcallback);
        }

        public bool GetObjectsUpDown(IDocumentOrFolder docorfolder, bool forcereadchildren, bool forcegetdzcollection, CallbackAction finalcallback)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            bool ret = false;

            try
            {
                // TS 18.02.13 war nur für Ordner
                // if (docorfolder != null && docorfolder.isFolder)
                // TS 04.03.16
                if (docorfolder.objectId.Equals(DataAdapter.Instance.DataCache.Objects(Workspace).Root.objectId))
                {
                    int test = 0;
                    test = test + 1;
                }
                // TS 04.03.16
                // if (docorfolder != null && !docorfolder.objectId.Equals(DataAdapter.Instance.DataCache.Objects(Workspace).Root.objectId) && docorfolder.objectId.Length > 0)
                // TS 24.03.16
                // if (docorfolder != null && docorfolder.objectId.Length > 0)
                bool stop = this.Workspace.ToString().Equals(CSEnumProfileWorkspace.workspace_deleted.ToString());
                stop = stop || this.Workspace.ToString().Equals(CSEnumProfileWorkspace.workspace_lastedited.ToString());
                // if (!this.Workspace.ToString().Equals(CSEnumProfileWorkspace.workspace_deleted.ToString()) && docorfolder != null && docorfolder.objectId.Length > 0)
                // TS 23.11.17 dokuemnte muessen das nicht machen
                stop = stop || docorfolder.isDocument;
                if (!stop && docorfolder != null && docorfolder.objectId.Length > 0)
                {
                    bool getparents = docorfolder.canGetObjectParents && !DataAdapter.Instance.DataCache.Objects(Workspace).IsStructRead(docorfolder.objectId);
                    bool getchilddocs = docorfolder.hasChildDocuments && docorfolder.ChildDocuments.Count == 0;
                    bool needsrefresh = docorfolder.needsCompleteRefresh;
                    docorfolder.needsForceUpdate = false;
                    // TS 15.02.16 die ersten child objekte immer einlesen wenn welche vorhanden und noch keine eingelesen
                    bool getchildfolders = docorfolder.hasChildFolders && docorfolder.ChildFolders.Count == 0;
                    // TS 18.03.16
                    if (getchildfolders)
                    {
                        if (DataAdapter.Instance.DataCache.Profile.Option_GetValueBoolean(CSEnumOptions.treealwaysloadall))
                        {
                            getchildfolders = false;
                        }
                        else
                        {
                            getchildfolders = DataAdapter.Instance.DataCache.Profile.Option_GetValueBoolean(CSEnumOptions.autoexpand_tree);
                        }
                    }

                    // TS 27.11.13 descendants lesen wenn eingeschaltet und:
                    // entweder bei getparents
                    // oder !getparents && istopnode && hasMoreItems
                    bool getdescendants = false;
                    bool option_loadAll = DataAdapter.Instance.DataCache.Profile.Option_GetValueBoolean(CSEnumOptions.treealwaysloadall);
                    CSEnumOptionPresets_querystructmode getdescendantsmode = CSEnumOptionPresets_querystructmode.querystructmode_base;

                    // ***************
                    // GET DESCENDANTS
                    // ***************
                    if (!docorfolder.isNotCreated && option_loadAll)
                    {
                        getdescendants = getparents;
                        if (!getdescendants)
                        {
                            bool istoplevel = docorfolder.objectId.Equals(DataAdapter.Instance.DataCache.Objects(this.Workspace).Folder_SelectedRoot.objectId);
                            if (istoplevel)
                            {
                                int desclevel = docorfolder.descendantsRead;
                                getdescendants = (desclevel >= 0);
                            }
                        }
                        if (getdescendants)
                        {
                            getdescendantsmode = CSEnumOptionPresets_querystructmode.querystructmode_all_levels_nodocs;
                        }
                    }

                    // getchildren
                    // TS 15.02.16 die ersten child objekte immer einlesen wenn welche vorhanden und noch keine eingelesen
                    //bool getchildren = getchilddocs || forcereadchildren;
                    bool getchildren = getchildfolders || getchilddocs || forcereadchildren || needsrefresh;

                    // getdzcollection
                    bool getdzcollection = false;
                    List<CServer.cmisObjectType> cmisdocuments = new List<CServer.cmisObjectType>();

                    // TS 23.03.16 nicht für den root ordner
                    // das muss nun hier rein weil oben die prüfung rausgenommen wurde wg. des aktenplans
                    // if (docorfolder.isFolder)
                    if (docorfolder.isFolder && !docorfolder.objectId.Equals(DataAdapter.Instance.DataCache.Objects(Workspace).Root.objectId))
                    {
                        // TS 01.02.16
                        //getdzcollection = getchilddocs || !docorfolder.isDZCollectionComplete;
                        // TS 03.02.16 forcegetdzcollection
                        //getdzcollection = getchilddocs || !docorfolder.isDZCollectionComplete || (docorfolder.hasChildDocuments && forcereadchildren);
                        getdzcollection = getchilddocs || !docorfolder.isDZCollectionComplete || (docorfolder.hasChildDocuments && forcereadchildren) || forcegetdzcollection || needsrefresh;
                    }
                    if (getdzcollection)
                        _CollectDZImagesRecursive(docorfolder, cmisdocuments);

                    // TS 20.02.13 z.b. für volltextsuche auch dokumente downloaden
                    bool getdownload = (docorfolder.isDocument && docorfolder.DownloadName.Length == 0);

                    // TS 15.02.16
                    // if (getparents || getdescendants || getchilddocs || forcereadchildren || getdownload)
                    if (getparents || getdescendants || getchildren || getdownload || needsrefresh)
                    {
                        string maxlistsize_folders = maxlistsize_folders = DataAdapter.Instance.DataCache.Profile.Option_GetValue(CServer.CSEnumOptions.querytreesize);
                        string maxlistsize_docs = DataAdapter.Instance.DataCache.Profile.Option_GetValue(CServer.CSEnumOptions.maxdownloads);

                        // TS 27.06.13 abholen von orderby presets // 01.10.13 oder auch von StructSorts aus dem Treeview gesetzt
                        string orderby = "";
                        List<CServer.CSOrderToken> givenorderby = new List<CSOrderToken>();
                        _Query_GetSortDataFilter(ref givenorderby, false);
                        foreach (CServer.CSOrderToken orderbytoken in givenorderby)
                        {
                            if (orderby.Length > 0)
                            {
                                // TS 01.10.13 muss komma sein statt blank !!
                                //orderby = orderby + " ";
                                orderby = orderby + ", ";
                            }
                            orderby = orderby + orderbytoken.propertyname + " " + orderbytoken.orderby;
                        }

                        // TS 21.01.16
                        // TS 22.01.16 doch immer beide lesen
                        //bool gettreeprops = getparents || getdescendants || getchilddocs || forcereadchildren;
                        //bool getlistprops = getparents;
                        //List<string> displayproperties = _Query_GetDisplayProperties(gettreeprops, getlistprops);
                        List<string> displayproperties = _Query_GetDisplayProperties(true, true);

                        string propertiesfilter = "";
                        foreach (string prop in displayproperties)
                        {
                            if (propertiesfilter.Length > 0) propertiesfilter = propertiesfilter + ",";
                            propertiesfilter = propertiesfilter + prop;
                        }

                        // TS 25.01.16
                        //bool getfulltextpreview = false;
                        bool getfulltextpreview = (ListDisplayMode == CSEnumOptionPresets_listdisplaymode.listdisplaymode_columnsandtext || ListDisplayMode == CSEnumOptionPresets_listdisplaymode.listdisplaymode_text);

                        DataAdapter.Instance.DataProvider.GetObjectsUpDown(docorfolder.CMISObject,
                                                            getparents, getchildren, getdzcollection, getfulltextpreview, getdescendantsmode, cmisdocuments,
                                                            propertiesfilter, orderby, int.Parse(maxlistsize_folders), int.Parse(maxlistsize_docs),
                                                            DataAdapter.Instance.DataCache.Rights.UserPrincipal, finalcallback);
                        ret = true;
                    }
                    // TS 23.11.17 die pruefung ist nun im server
                    // else if (getdzcollection && !docorfolder.isNotCreated)
                    else if (getdzcollection)
                        DataAdapter.Instance.DataProvider.GetDZCollection(docorfolder.CMISObject, cmisdocuments, DataAdapter.Instance.DataCache.Rights.UserPrincipal, finalcallback);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
            return ret;
        }

        #endregion GetObjectsUpDown

        // ==================================================================

        #region GetTemplateDefinitions (cmd_GetTemplateDefinitions)

        public DelegateCommand cmd_GetTemplateDefinitions { get { return new DelegateCommand(GetTemplateDefinitions, _CanGetTemplateDefinitions); } }

        public void GetTemplateDefinitions()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanGetTemplateDefinitions)
                {
                    string customPath = DataAdapter.Instance.DataCache.Profile.Option_GetValue(CSEnumOptions.custommfppath);
                    DataAdapter.Instance.DataProvider.GetTemplateDefinitions(DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_default).Root.CMISObject,
                    CSEnumProfileWorkspace.workspace_clipboard, DataAdapter.Instance.DataCache.Rights.UserPrincipal);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        public bool CanGetTemplateDefinitions { get { return _CanGetTemplateDefinitions(); } }

        private bool _CanGetTemplateDefinitions()
        {
            return DataAdapter.Instance.DataCache.ApplicationFullyInit;
        }

        #endregion GetTemplateDefinitions (cmd_GetTemplateDefinitions)

        // ==================================================================


        // ==================================================================

        #region HookDel (cmd_HookDel)

        public DelegateCommand cmd_HookDel { get { return new DelegateCommand(HookDel, _CanHookDel); } }

        public void HookDel()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanHookDel)
                {
                    // Create empty Hook
                    CSProfileCmisObject hook = DataAdapter.Instance.DataCache.Profile.UserProfile.objects.hooked;
                    hook.repository = "";
                    hook.objectid = "";
                    hook.comment = "";

                    // Save Hook
                    Profile_WriteProfile();
                    DataAdapter.Instance.DataCache.Profile.UserProfile.objects.hooked = hook;
                    DataAdapter.Instance.InformObservers();
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public bool CanHookDel { get { return _CanHookDel(); } }

        private bool _CanHookDel()
        {
            return DataAdapter.Instance.DataCache.ApplicationFullyInit && DataAdapter.Instance.DataCache.Profile.UserProfile.objects.hooked.objectid != "";
        }

        #endregion HookDel (cmd_HookDel)

        // ==================================================================

        #region HookGet (cmd_HookGet)

        public DelegateCommand cmd_HookGet { get { return new DelegateCommand(HookGet, _CanHookGet); } }

        public void HookGet_Hotkey() { HookGet(); }
        public void HookGet()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanHookGet)
                {
                    bool changedRep = false;
                    CSProfileCmisObject hook = DataAdapter.Instance.DataCache.Profile.UserProfile.objects.hooked;
                    ClearCache(false, null);

                    // Jump to the right repository, if we aren't in the right one
                    if (!hook.repository.Equals(DataAdapter.Instance.DataCache.Rights.ChoosenRepository.repositoryId))
                    {
                        CSRootNode rootnode = DataAdapter.Instance.DataCache.RootNodes.GetNodeById(hook.repository);
                        if (rootnode != null)
                        {
                            changedRep = true;
                            AppManager.StartupQueryObjectid = hook.objectid;
                            AppManager.ChooseRootNode(rootnode, false);
                        }
                    }
                    if (!changedRep)
                    {
                        QuerySingleObjectById(hook.repository, hook.objectid, false, true, CSEnumProfileWorkspace.workspace_default, null);
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public bool CanHookGet { get { return _CanHookGet(); } }

        private bool _CanHookGet()
        {
            return DataAdapter.Instance.DataCache.ApplicationFullyInit && DataAdapter.Instance.DataCache.Profile.UserProfile.objects.hooked.objectid != "";
        }

        #endregion HookGet (cmd_HookGet)

        // ==================================================================

        #region HookSet (cmd_HookSet)

        public DelegateCommand cmd_HookSet { get { return new DelegateCommand(HookSet, _CanHookSet); } }

        public void HookSet()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanHookSet)
                {
                    IDocumentOrFolder current = DataAdapter.Instance.DataCache.Objects(Workspace).Object_Selected;

                    // Create new Hook
                    CSProfileCmisObject hook = DataAdapter.Instance.DataCache.Profile.UserProfile.objects.hooked;
                    hook.repository = current.RepositoryId;
                    hook.objectid = current.objectId;
                    hook.objecttypeid = current.objectTypeId;
                    hook.editable = true;
                    hook.comment = System.DateTime.Now.ToString();

                    // Save Hook
                    Profile_WriteProfile();
                    DataAdapter.Instance.DataCache.Profile.UserProfile.objects.hooked = hook;
                    DataAdapter.Instance.InformObservers();
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public bool CanHookSet { get { return _CanHookSet(); } }

        private bool _CanHookSet()
        {
            return DataAdapter.Instance.DataCache.ApplicationFullyInit
                && DataAdapter.Instance.DataCache.Objects(Workspace).Object_Selected.objectId != ""
                && DataAdapter.Instance.DataCache.Objects(Workspace).Object_Selected.objectId != DataAdapter.Instance.DataCache.Objects(Workspace).Root.objectId
                && DataAdapter.Instance.DataCache.Objects(Workspace).Object_Selected.isFolder;
        }

        #endregion HookSet (cmd_HookSet)

        // ==================================================================

        #region Local Connector

        // ==================================================================

        #region LC_Init

        public DelegateCommand cmd_LC_Init { get { return new DelegateCommand(LC_Init, _CanLC_Init); } }

        // TS 21.07.15
        //public void LC_Init() { LC_Init(false); }
        public void LC_Init(object autostart) { LC_Init(autostart, false); }

        public void LC_Init(object autostart, bool pushtolc)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanLC_Init)
                {
                    string lcexecutefile = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.LC_Command];
                    string lcparamfilename = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.LC_Param];
                    CSLCRequest lcrequest = CreateLCRequest();

                    // TS 22.07.15
                    bool writeautostart = false;
                    if (autostart != null)
                        writeautostart = (bool)autostart;
                    lcrequest.install_autostart = writeautostart;

                    // TS 21.07.15
                    //bool pushtolc = false;
                    DataAdapter.Instance.DataProvider.LC_CreateCommandFile(CSEnumCommands.cmd_LC_Init, lcparamfilename, lcrequest, pushtolc, DataAdapter.Instance.DataCache.Rights.UserPrincipal);
                    // der call hier geht immer über clickonce damit die exe gestartet wird
                    // TS 21.07.15 das stimmt nicht mehr, dank lc_pending kann er auch anders ...
                    // if (lcexecutefile != null && lcexecutefile.Length > 0 && lcparamfilename != null && lcparamfilename.Length > 0)
                    if (!pushtolc && lcexecutefile != null && lcexecutefile.Length > 0 && lcparamfilename != null && lcparamfilename.Length > 0)
                    {
                        LC_ExecuteCall(lcexecutefile, lcparamfilename);

                        // TS 28.05.18
                        if (AppManager.RCListener != null)
                        {
                            AppManager.RCListener.responseOnLCInformationRequested = true;
                        }

                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        private bool CanLC_Init { get { return _CanLC_Init(); } }

        private bool _CanLC_Init()
        {
            // TS 11.02.14
            // return DataAdapter.Instance.DataCache.ApplicationFullyInit && DataAdapter.Instance.DataCache.IsLocalConnectorAvailable;
            return DataAdapter.Instance.DataCache.ApplicationFullyInit && LC_IsAvailable;
        }

        public void LC_SendDummyCommand()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanLC_Init)
                {
                    string lcexecutefile = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.LC_Command];
                    string lcparamfilename = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.LC_Param];
                    CSLCRequest lcrequest = CreateLCRequest();
                    bool pushtolc = true;
                    DataAdapter.Instance.DataProvider.LC_CreateCommandFile(CSEnumCommands.cmd_LC_CheckAvail, lcparamfilename, lcrequest, pushtolc, DataAdapter.Instance.DataCache.Rights.UserPrincipal);                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }


        // ==================================================================

        #region LC_DocumentCreateWord (cmd_LC_DocumentCreateWord)

        public DelegateCommand cmd_LC_DocumentCreateWord { get { return new DelegateCommand(LC_DocumentCreateWord, _CanLC_DocumentCreateWord); } }

        public void LC_DocumentCreateWord()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanLC_DocumentCreateWord)
                {
                    string lcexecutefile = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.LC_Command];
                    string lcparamfilename = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.LC_Param];
                    CSLCRequest lcrequest = CreateLCRequest();

                    // TS 18.03.16 vorlagen
                    CSOption option_officetemplatemode = DataAdapter.Instance.DataCache.Profile.Option_GetOption(CSEnumOptions.officetemplatemode);
                    //CSEnumOptionPresets_officetemplatemode officetemplatemode = CSEnumOptionPresets_officetemplatemode.officetemplatemode_none;
                    CSEnumOptionPresets_officetemplatemode currentval = DataAdapter.Instance.DataCache.Profile.Option_GetEnumValue<CSEnumOptionPresets_officetemplatemode>(option_officetemplatemode);

                    // TS 18.03.16 vorlagen
                    lcrequest.proc_param_showdialog = 0;
                    if (currentval.Equals(CSEnumOptionPresets_officetemplatemode.officetemplatemode_ecm))
                    {
                        //getdescendantsmode = CSEnumOptionPresets_querystructmode.querystructmode_two_levels_nodocs;
                        // ECM noch nicht ausprogrammiert, kommmt dann hierhin
                        // Angedachte Logik:
                        // ECM Vorlage wählen
                        //   wenn keine gewählt dann zusätzlich MS Vorlage wählen
                        // zunächst auf MS setzen
                        lcrequest.proc_param_showdialog = 1;
                    }
                    else if (currentval.Equals(CSEnumOptionPresets_officetemplatemode.officetemplatemode_ms))
                    {
                        lcrequest.proc_param_showdialog = 2;
                    }

                    bool pushtolc = LC_IsPushEnabled;

                    DataAdapter.Instance.DataProvider.LC_CreateCommandFile(CSEnumCommands.cmd_LC_DocumentCreateWord, lcparamfilename, lcrequest, pushtolc, DataAdapter.Instance.DataCache.Rights.UserPrincipal);
                    // LC nur rufen nur wenn nicht bereits beim serveraufruf gepushed wurde
                    if (!pushtolc && lcexecutefile != null && lcexecutefile.Length > 0 && lcparamfilename != null && lcparamfilename.Length > 0)
                    {
                        LC_ExecuteCall(lcexecutefile, lcparamfilename);
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public bool CanLC_DocumentCreateWord { get { return _CanLC_DocumentCreateWord(); } }

        private bool _CanLC_DocumentCreateWord()
        {
            IDocumentOrFolder selObj = DataAdapter.Instance.DataCache.Objects(this.Workspace).Folder_Selected;

            return DataAdapter.Instance.DataCache.ApplicationFullyInit && LC_IsAvailable && LC_IsModulAvailable(CSEnumInformationId.LC_Modul_Word, CSEnumInformationId.LC_Modul_WordParam)
                && DataAdapter.Instance.DataCache.Objects(Workspace).Folder_Selected.objectId != ""
                && DataAdapter.Instance.DataCache.Objects(Workspace).Folder_Selected.objectId != DataAdapter.Instance.DataCache.Objects(Workspace).Root.objectId
                && CanCreateDocuments && selObj.hasCreatePermission();
        }

        #endregion LC_DocumentCreateWord (cmd_LC_DocumentCreateWord)

        // ==================================================================

        #region LC_DocumentCreateExcel (cmd_LC_DocumentCreateExcel)

        public DelegateCommand cmd_LC_DocumentCreateExcel { get { return new DelegateCommand(LC_DocumentCreateExcel, _CanLC_DocumentCreateExcel); } }

        public void LC_DocumentCreateExcel()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanLC_DocumentCreateExcel)
                {
                    string lcexecutefile = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.LC_Command];
                    string lcparamfilename = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.LC_Param];
                    CSLCRequest lcrequest = CreateLCRequest();

                    // TS 18.03.16 vorlagen
                    CSOption option_officetemplatemode = DataAdapter.Instance.DataCache.Profile.Option_GetOption(CSEnumOptions.officetemplatemode);
                    //CSEnumOptionPresets_officetemplatemode officetemplatemode = CSEnumOptionPresets_officetemplatemode.officetemplatemode_none;
                    CSEnumOptionPresets_officetemplatemode currentval = DataAdapter.Instance.DataCache.Profile.Option_GetEnumValue<CSEnumOptionPresets_officetemplatemode>(option_officetemplatemode);

                    // TS 18.03.16 vorlagen
                    lcrequest.proc_param_showdialog = 0;
                    if (currentval.Equals(CSEnumOptionPresets_officetemplatemode.officetemplatemode_ecm))
                    {
                        //getdescendantsmode = CSEnumOptionPresets_querystructmode.querystructmode_two_levels_nodocs;
                        // ECM noch nicht ausprogrammiert, kommmt dann hierhin
                        // Angedachte Logik:
                        // ECM Vorlage wählen
                        //   wenn keine gewählt dann zusätzlich MS Vorlage wählen
                        // zunächst auf MS setzen
                        lcrequest.proc_param_showdialog = 1;
                    }
                    else if (currentval.Equals(CSEnumOptionPresets_officetemplatemode.officetemplatemode_ms))
                    {
                        lcrequest.proc_param_showdialog = 2;
                    }

                    bool pushtolc = LC_IsPushEnabled;

                    DataAdapter.Instance.DataProvider.LC_CreateCommandFile(CSEnumCommands.cmd_LC_DocumentCreateExcel, lcparamfilename, lcrequest, pushtolc, DataAdapter.Instance.DataCache.Rights.UserPrincipal);
                    // LC nur rufen nur wenn nicht bereits beim serveraufruf gepushed wurde
                    if (!pushtolc && lcexecutefile != null && lcexecutefile.Length > 0 && lcparamfilename != null && lcparamfilename.Length > 0)
                    {
                        LC_ExecuteCall(lcexecutefile, lcparamfilename);
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public bool CanLC_DocumentCreateExcel { get { return _CanLC_DocumentCreateExcel(); } }

        private bool _CanLC_DocumentCreateExcel()
        {
            IDocumentOrFolder selObj = DataAdapter.Instance.DataCache.Objects(this.Workspace).Folder_Selected;

            return DataAdapter.Instance.DataCache.ApplicationFullyInit && LC_IsAvailable && LC_IsModulAvailable(CSEnumInformationId.LC_Modul_Excel, CSEnumInformationId.LC_Modul_ExcelParam)
                && DataAdapter.Instance.DataCache.Objects(Workspace).Folder_Selected.objectId != ""
                && DataAdapter.Instance.DataCache.Objects(Workspace).Folder_Selected.objectId != DataAdapter.Instance.DataCache.Objects(Workspace).Root.objectId
                && CanCreateDocuments && selObj.hasCreatePermission();
        }

        #endregion LC_DocumentCreateExcel (cmd_LC_DocumentCreateExcel)

        // ==================================================================
        //MRWORK
        #region cmd_LC_DocumentMailMerge (cmd_LC_DocumentMailMerge)

        public DelegateCommand cmd_LC_DocumentMailMerge { get { return new DelegateCommand(LC_DocumentMailMerge, _CanLC_DocumentMailMerge); } }

        public void LC_DocumentMailMerge()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanLC_DocumentMailMerge)
                {
                    string lcexecutefile = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.LC_Command];
                    string lcparamfilename = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.LC_Param];
                    CSLCRequest lcrequest = CreateLCRequest();

                    // TS 18.03.16 vorlagen
                    CSOption option_officetemplatemode = DataAdapter.Instance.DataCache.Profile.Option_GetOption(CSEnumOptions.officetemplatemode);
                    //CSEnumOptionPresets_officetemplatemode officetemplatemode = CSEnumOptionPresets_officetemplatemode.officetemplatemode_none;
                    CSEnumOptionPresets_officetemplatemode currentval = DataAdapter.Instance.DataCache.Profile.Option_GetEnumValue<CSEnumOptionPresets_officetemplatemode>(option_officetemplatemode);

                    //string[] objids=new string[];
                    

                    List<IDocumentOrFolder> targetList = DataAdapter.Instance.DataCache.Objects(Workspace).ObjectList.ToList();
                    List<string> targetObjIDs = new List<string>();

                    // TS 16.03.17
                    //List <IDocumentOrFolder> targetList = DataAdapter.Instance.DataCache.Objects(Workspace).ObjectList.ToList();
                    //foreach (IDocumentOrFolder doc in targetList)
                    //{
                    //    if (doc.isAddressObject && doc.isQueryResult)
                    //    {
                    //        targetObjIDs.Add(doc.objectId);
                    //    }
                    //}
                    foreach (IDocumentOrFolder obj in DataAdapter.Instance.DataCache.Objects(Workspace).ObjectList_QueryResults)
                    {
                        if (obj.isAddressObject && obj.isUserSelected)
                        {
                            targetObjIDs.Add(obj.objectId);
                        }
                    }
                    lcrequest.proc_param_objectids = targetObjIDs.ToArray();

                    // TS 18.03.16 vorlagen
                    lcrequest.proc_param_showdialog = 0;
                    if (currentval.Equals(CSEnumOptionPresets_officetemplatemode.officetemplatemode_ecm))
                    {
                        //getdescendantsmode = CSEnumOptionPresets_querystructmode.querystructmode_two_levels_nodocs;
                        // ECM noch nicht ausprogrammiert, kommmt dann hierhin
                        // Angedachte Logik:
                        // ECM Vorlage wählen
                        //   wenn keine gewählt dann zusätzlich MS Vorlage wählen
                        // zunächst auf MS setzen
                        lcrequest.proc_param_showdialog = 1;
                    }
                    else if (currentval.Equals(CSEnumOptionPresets_officetemplatemode.officetemplatemode_ms))
                    {
                        lcrequest.proc_param_showdialog = 2;
                    }

                    bool pushtolc = LC_IsPushEnabled;

                    DataAdapter.Instance.DataProvider.LC_CreateCommandFile(CSEnumCommands.cmd_LC_DocumentMailMerge, lcparamfilename, lcrequest, pushtolc, DataAdapter.Instance.DataCache.Rights.UserPrincipal);
                    // LC nur rufen nur wenn nicht bereits beim serveraufruf gepushed wurde
                    if (!pushtolc && lcexecutefile != null && lcexecutefile.Length > 0 && lcparamfilename != null && lcparamfilename.Length > 0)
                    {
                        LC_ExecuteCall(lcexecutefile, lcparamfilename);
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public bool CanLC_DocumentMailMerge { get { return _CanLC_DocumentMailMerge(); } }

        private bool _CanLC_DocumentMailMerge()
        {
            if(Workspace != CSEnumProfileWorkspace.workspace_searchoverall)
            {
                return DataAdapter.Instance.DataCache.ApplicationFullyInit && LC_IsAvailable && LC_IsModulAvailable(CSEnumInformationId.LC_Modul_Word, CSEnumInformationId.LC_Modul_WordParam)
                    && DataAdapter.Instance.DataCache.Objects(Workspace).Folder_Selected.objectId != ""
                    && DataAdapter.Instance.DataCache.Objects(Workspace).Folder_Selected.objectId != DataAdapter.Instance.DataCache.Objects(Workspace).Root.objectId;
            }else
            {
                return DataAdapter.Instance.DataCache.ApplicationFullyInit && LC_IsAvailable && LC_IsModulAvailable(CSEnumInformationId.LC_Modul_Word, CSEnumInformationId.LC_Modul_WordParam)
                    && DataAdapter.Instance.DataCache.Objects(Workspace).ObjectList.Count > 0;
            }
        }

        #endregion LC_DocumentCreateWord (cmd_LC_DocumentCreateWord)

        // ==================================================================
        #region LC_CreateDocVersionFrom (cmd_LC_CreateDocIVersionFrom und cmd_LC_CreateDocOVersionFrom)

        public DelegateCommand cmd_LC_CreateDocIVersionFrom { get { return new DelegateCommand(LC_CreateDocIVersionFrom, _CanLC_CreateDocVersionFrom); } }
        public DelegateCommand cmd_LC_CreateDocOVersionFrom { get { return new DelegateCommand(LC_CreateDocOVersionFrom, _CanLC_CreateDocVersionFrom); } }

        public void LC_CreateDocIVersionFrom()
        {
            LC_CreateDocVersionFrom(CSEnumDocumentCopyTypes.COPYTYPE_CONTENT);
        }

        public void LC_CreateDocOVersionFrom()
        {
            LC_CreateDocVersionFrom(CSEnumDocumentCopyTypes.COPYTYPE_ORIGINAL);
        }

        private void LC_CreateDocVersionFrom(CSEnumDocumentCopyTypes doccopytype)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanLC_CreateDocVersionFrom)
                {
                    string lcexecutefile = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.LC_Command];
                    string lcparamfilename = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.LC_Param];
                    CSLCRequest lcrequest = CreateLCRequest();
                    lcrequest.proc_repositoryid = DataAdapter.Instance.DataCache.Objects(this.Workspace).Document_Selected.RepositoryId;
                    lcrequest.proc_parentobjectid = DataAdapter.Instance.DataCache.Objects(this.Workspace).Document_Selected.objectId;
                    lcrequest.proc_param_application = LC_GetDocEditApplication(DataAdapter.Instance.DataCache.Objects(this.Workspace).Document_Selected.DownloadName);
                    lcrequest.proc_param_doccopytype = doccopytype;
                    lcrequest.proc_param_doccopytypeSpecified = true;

                    bool pushtolc = LC_IsPushEnabled;

                    if (doccopytype.ToString().Equals(CSEnumDocumentCopyTypes.COPYTYPE_ORIGINAL.ToString()))
                        DataAdapter.Instance.DataProvider.LC_CreateCommandFile(CSEnumCommands.cmd_LC_CreateDocOVersionFrom, lcparamfilename, lcrequest, pushtolc, DataAdapter.Instance.DataCache.Rights.UserPrincipal);
                    else if (doccopytype.ToString().Equals(CSEnumDocumentCopyTypes.COPYTYPE_CONTENT.ToString()))
                        DataAdapter.Instance.DataProvider.LC_CreateCommandFile(CSEnumCommands.cmd_LC_CreateDocIVersionFrom, lcparamfilename, lcrequest, pushtolc, DataAdapter.Instance.DataCache.Rights.UserPrincipal);
                    else
                        lcexecutefile = "";
                    // LC nur rufen nur wenn nicht bereits beim serveraufruf gepushed wurde
                    if (!pushtolc && lcexecutefile != null && lcexecutefile.Length > 0 && lcparamfilename != null && lcparamfilename.Length > 0)
                    {
                        LC_ExecuteCall(lcexecutefile, lcparamfilename);
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public bool CanLC_CreateDocVersionFrom { get { return _CanLC_CreateDocVersionFrom(); } }

        private bool _CanLC_CreateDocVersionFrom()
        {
            return DataAdapter.Instance.DataCache.ApplicationFullyInit && LC_IsAvailable
                && DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected.objectId != ""
                && DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected.objectId != DataAdapter.Instance.DataCache.Objects(Workspace).Root.objectId
                && LC_GetDocEditApplication(DataAdapter.Instance.DataCache.Objects(this.Workspace).Document_Selected.DownloadName).Length > 0;
        }

        #endregion LC_CreateDocVersionFrom (cmd_LC_CreateDocIVersionFrom und cmd_LC_CreateDocOVersionFrom)

        // ==================================================================
        #region LC_DocumentSign (cmd_LC_DocumentSign)
        public DelegateCommand cmd_LC_DocumentSign { get { return new DelegateCommand(LC_DocumentSign, _CanLC_DocumentSign); } }


        public void LC_DocumentSign()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanLC_DocumentSign)
                {
                    string lcexecutefile = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.LC_Command];
                    string lcparamfilename = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.LC_Param];
                    CSLCRequest lcrequest = CreateLCRequest();
                    lcrequest.proc_repositoryid = DataAdapter.Instance.DataCache.Objects(this.Workspace).Document_Selected.RepositoryId;
                    
                    DocumentOrFolder dof = (DocumentOrFolder)DataAdapter.Instance.DataCache.Objects(this.Workspace).Document_Selected;
                    string typeid=dof.objectTypeId;
                    string did = "";
                    string pid = "";
                    if(dof.isDocument)
                    {
                        did = dof.objectId;
                        pid = dof.parentId;
                    }
                    string[] dids = new string[1];
                    dids[0] = did;
                    lcrequest.proc_param_objectids = dids;
                    lcrequest.proc_parentobjectid = pid;

                    

                    bool pushtolc = LC_IsPushEnabled;

                    DataAdapter.Instance.DataProvider.LC_CreateCommandFile(CSEnumCommands.cmd_LC_DocumentSign, lcparamfilename, lcrequest, pushtolc, DataAdapter.Instance.DataCache.Rights.UserPrincipal);
                    
                    // LC nur rufen nur wenn nicht bereits beim serveraufruf gepushed wurde
                    if (!pushtolc && lcexecutefile != null && lcexecutefile.Length > 0 && lcparamfilename != null && lcparamfilename.Length > 0)
                    {
                        LC_ExecuteCall(lcexecutefile, lcparamfilename);
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public bool CanLC_DocumentSign { get { return _CanLC_DocumentSign(); } }

        private bool _CanLC_DocumentSign()
        {
            IDocumentOrFolder selObj = DataAdapter.Instance.DataCache.Objects(this.Workspace).Folder_Selected;

            return DataAdapter.Instance.DataCache.ApplicationFullyInit && LC_IsAvailable
                && DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected.objectId != ""
                && DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected.objectId != DataAdapter.Instance.DataCache.Objects(Workspace).Root.objectId
                && DataAdapter.Instance.DataCache.Objects(this.Workspace).Document_Selected.DownloadName.Length > 0
                && selObj.hasCreatePermission();
                //&& LC_GetDocEditApplication(DataAdapter.Instance.DataCache.Objects(this.Workspace).Document_Selected.DownloadName).Length > 0;
        }

        #endregion LC_CreateDocVersionFrom (cmd_LC_CreateDocIVersionFrom und cmd_LC_CreateDocOVersionFrom)
        // ==================================================================
        // mail erstellen (mail wird über server in postfach erstellt, lc zeigt es dann an über LC_MailShow)

        #region LC_MailCreate (cmd_LC_MailCreate)

        public DelegateCommand cmd_LC_MailCreate { get { return new DelegateCommand(LC_MailCreate, _CanLC_MailCreate); } }

        public void LC_MailCreate_Hotkey() { LC_MailCreate(); }
        public void LC_MailCreate()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanLC_MailCreate)
                {
                    List<cmisObjectType> docstoappend = new List<cmisObjectType>();
                    foreach (IDocumentOrFolder obj in DataAdapter.Instance.DataCache.Objects(Workspace).Objects_Selected)
                    {
                        if (obj.structLevel >= Statics.Constants.STRUCTLEVEL_09_DOKLOG)
                        {
                            docstoappend.Add(obj.CMISObject);
                        }
                    }
                    List<string> recipient_to = new List<string>();
                    List<string> recipient_cc = new List<string>();
                    List<string> recipient_bcc = new List<string>();
                    string from="";
                    string subject="";
                    string content="";

                    LC_MailCreateGetMailFieldValues(ref recipient_to, ref recipient_cc, ref recipient_bcc, ref from, ref subject, ref content);
                    LC_MailCreate(recipient_to, null, null, null, null, content, null, docstoappend, false, null);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }
        
        public DelegateCommand cmd_LC_MailCreateWithSample { get { return new DelegateCommand(LC_MailCreateWithSample, _CanLC_MailCreate); } }

        public void LC_MailCreateWithSample_Hotkey() { LC_MailCreateWithSample(); }
        public void LC_MailCreateWithSample()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanLC_MailCreate)
                {
                    List<cmisObjectType> docstoappend = new List<cmisObjectType>();
                    foreach (IDocumentOrFolder obj in DataAdapter.Instance.DataCache.Objects(Workspace).Objects_Selected)
                    {
                        if (obj.structLevel >= Statics.Constants.STRUCTLEVEL_09_DOKLOG)
                        {
                            docstoappend.Add(obj.CMISObject);
                        }
                    }
                    List<string> recipient_to = new List<string>();
                    List<string> recipient_cc = new List<string>();
                    List<string> recipient_bcc = new List<string>();
                    string from = "";
                    string subject = "";
                    string content = "";

                    LC_MailCreateGetMailFieldValues(ref recipient_to, ref recipient_cc, ref recipient_bcc, ref from, ref subject, ref content);
                    LC_MailCreate(recipient_to, null, null, null, null, content, null, docstoappend, false, null, true);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }


        //MRWORK
        public void LC_MailCreateGetMailFieldValues(ref List<string> recipient_to, ref List<string> recipient_cc, ref List<string> recipient_bcc, ref string from, ref string subject, ref string content)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            
            List<string> values = null;
            
            
            try
            {
                IDocumentOrFolder obj = DataAdapter.Instance.DataCache.Objects(Workspace).Object_Selected;
                if(obj!=null)
                {
                    //if (!obj.isAddressObject)
                    if (obj.structLevel == Statics.Constants.STRUCTLEVEL_09_DOKLOG) obj = obj.Parent;
                    if (obj.isAddressObject)
                    {
                        //emailadresse 
                        string mailto_sw = "E-MAIL" + Constants.MULTIVALUEBINDER_SEPARATOR; //muss noch geändert werden!
                        if (obj.isOldADRModel)
                        {
                            values = obj.GetCMISPropertyAllValues(CSEnumCmisProperties.V_ADR_15);
                        }
                        else
                        {
                            values = obj.GetCMISPropertyAllValues(CSEnumCmisProperties.VORGANG_ADR_15);
                        }
                        foreach (string value in values)
                        {
                            if (value.ToUpper().StartsWith(mailto_sw))
                            {
                                char splitter = (Constants.MULTIVALUEBINDER_SEPARATOR.ToCharArray())[0];
                                string[] mailto_parts = value.Split(splitter);
                                recipient_to.Add(mailto_parts[1]);
                            }
                            
                        }
                        //anrede für body
                        if (obj.isOldADRModel)
                        {
                            values = obj.GetCMISPropertyAllValues(CSEnumCmisProperties.V_ADR_05);
                        }
                        else
                        {
                            values = obj.GetCMISPropertyAllValues(CSEnumCmisProperties.VORGANG_ADR_05);
                        }
                        if (values.Count > 0) content = values[0];
                    }
            }
                
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }
        // TS 30.09.15 einstieg über callback wg. login dialog
        private void LC_MailCreate(List<object> parameters)
        {
            if (parameters.GetType().Name.StartsWith("List"))
            {
                List<object> paramlist = (List<object>)parameters;
                if (paramlist.Count >= 9)
                {
                    List<string> recipient_to = null;
                    if (paramlist[0] != null) recipient_to = (List<string>)paramlist[0];
                    List<string> recipient_cc = null;
                    if (paramlist[1] != null) recipient_cc = (List<string>)paramlist[1];
                    List<string> recipient_bcc = null;
                    if (paramlist[2] != null) recipient_bcc = (List<string>)paramlist[2];
                    string from = null;
                    if (paramlist[3] != null) from = (string)paramlist[3];
                    string subject = null;
                    if (paramlist[4] != null) subject = (string)paramlist[4];
                    string content = null;
                    if (paramlist[5] != null) content = (string)paramlist[5];
                    string targetfoldername = null;
                    if (paramlist[6] != null) targetfoldername = (string)paramlist[6];
                    List<cmisObjectType> docstoappend = null;
                    if (paramlist[7] != null) docstoappend = (List<cmisObjectType>)paramlist[7];
                    bool askcredentials = false;
                    if (paramlist[8] != null) askcredentials = (bool)paramlist[8];
                    string askedpassword = null;
                    if (paramlist.Count > 9 && paramlist[9] != null) askedpassword = (string)paramlist[9];

                    LC_MailCreate(recipient_to, recipient_cc, recipient_bcc, from, subject, content, targetfoldername, docstoappend, askcredentials, askedpassword);
                }
            }
        }

        public void LC_MailCreate(List<string> recipient_to, List<string> recipient_cc, List<string> recipient_bcc, string from, string subject, string content, string targetfoldername,
            List<cmisObjectType> docstoappend, bool askcredentials, string askedpassword)
        {
            LC_MailCreate(recipient_to, recipient_cc, recipient_bcc, from, subject, content, targetfoldername, docstoappend, askcredentials, askedpassword, false);
        }

        public void LC_MailCreate(List<string> recipient_to, List<string> recipient_cc, List<string> recipient_bcc, string from, string subject, string content, string targetfoldername,
            List<cmisObjectType> docstoappend, bool askcredentials, string askedpassword, bool showsamplewindow)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                // TS 30.09.15
                List<object> callbackparams = new List<object>();
                callbackparams.Add(recipient_to);
                callbackparams.Add(recipient_cc);
                callbackparams.Add(recipient_bcc);
                callbackparams.Add(from);
                callbackparams.Add(subject);
                callbackparams.Add(content);
                callbackparams.Add(targetfoldername);
                callbackparams.Add(docstoappend);

                if (askcredentials)
                {
                    // TS 30.09.15
                    if (DataAdapter.Instance.DataCache.ResponseStatus.returncode.Equals(Statics.Constants.SERVERERRORCODE_MAILCONNECT_FAIL)
                        ||
                        DataAdapter.Instance.DataCache.ResponseStatus.returncode.Equals(Statics.Constants.SERVERERRORCODE_MAILLISTENER_FAIL))
                    {
                        callbackparams.Add(false);
                        View.Dialogs.dlgAskPW askpw = new View.Dialogs.dlgAskPW(LC_MailCreate, callbackparams, null);
                        DialogHandler.Show_Dialog(askpw);
                    }
                }
                else
                {
                    // TS 30.09.15 callback falls nicht an mailserver angemeldet werden konnte dann nochmal hier aufrufen mit askcredentials = true
                    callbackparams.Add(true);
                    CallbackAction callbackiffailed = new CallbackAction(LC_MailCreate, callbackparams);

                    // TS 30.09.15 neuen principal verwenden wg. kennwortabfrage
                    // TS 11.01.17 principal erweitert, das mailserverpw ist nun als extra feld vorhanden
                    //CSUserPrincipal localprincipal = DataCache_Rights.CopyPrincipal(DataAdapter.Instance.DataCache.Rights.UserPrincipal);                    
                    //if (askedpassword != null && askedpassword.Length > 0) localprincipal.password = askedpassword;
                    CSUserPrincipal principal = DataAdapter.Instance.DataCache.Rights.UserPrincipal;
                    if (askedpassword != null && askedpassword.Length > 0) principal.mailserverpassword = askedpassword;

                    bool pushtolc = LC_IsPushEnabled;
                    string mailserverclientcall = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.LC_MailServerClientCall];
                    if (mailserverclientcall == null) mailserverclientcall = "";

                    CSOption showcopytypeoption = DataAdapter.Instance.DataCache.Profile.Option_GetOption(CServer.CSEnumOptions.showcopytypes);
                    CSEnumOptionPresets_showcopytypes showcopytype = DataAdapter.Instance.DataCache.Profile.Option_GetEnumValue<CSEnumOptionPresets_showcopytypes>(showcopytypeoption);

                    // TS 14.12.15 alte mailid aus infocache vorher löschen
                    CSInformation[] infoarray = new CSInformation[1];
                    infoarray[0] = new CSInformation();
                    infoarray[0].informationid = CSEnumInformationId.MailId_Created;
                    infoarray[0].informationidSpecified = true;
                    infoarray[0].informationvalue = "";
                    DataAdapter.Instance.DataCache.ApplyData(infoarray);

                    if (mailserverclientcall.Length == 0)
                        DataAdapter.Instance.DataProvider.CreateMail(docstoappend, recipient_to, recipient_cc, recipient_bcc, from, subject, content, targetfoldername, showcopytype,
                            principal, new CallbackAction(LC_MailCreate_Done_SendLC, pushtolc, showsamplewindow), callbackiffailed);
                    else
                        DataAdapter.Instance.DataProvider.CreateMail(docstoappend, recipient_to, recipient_cc, recipient_bcc, from, subject, content, targetfoldername, showcopytype,
                            principal, new CallbackAction(LC_MailCreate_Done_SendClientCall, mailserverclientcall), callbackiffailed);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public bool CanLC_MailCreate { get { return _CanLC_MailCreate(); } }

        private bool _CanLC_MailCreate()
        {
            bool ret = DataAdapter.Instance.DataCache.ApplicationFullyInit;
            if (ret)
            {
                bool anyfound = false;
                bool anyinvalid = false;
                foreach (IDocumentOrFolder obj in DataAdapter.Instance.DataCache.Objects(Workspace).Objects_Selected)
                {
                    if (obj.structLevel >= Statics.Constants.STRUCTLEVEL_09_DOKLOG || obj.structLevel == Statics.Constants.STRUCTLEVEL_08_VORGANG || obj.structLevel == Statics.Constants.STRUCTLEVEL_07_AKTE)
                        anyfound = true;
                    else
                    {
                        anyinvalid = true;
                        break;
                    }
                }
                
                ret = anyfound && !anyinvalid;
            }
            return ret;
        }

        private bool _CanLC_MailCreate_IsLCAvail()
        {
            return CanLC_MailCreate && LC_IsAvailable && LC_IsModulAvailable(CSEnumInformationId.LC_Modul_Outlook, CSEnumInformationId.LC_Modul_OutlookParam);
        }

        private void LC_MailCreate_Done_SendLC(object pushtolc_obj, object showsamplewindow_obj)
        {
            try
            {
                bool pushtolc = (bool)pushtolc_obj;
                bool showsamplewindow = (bool)showsamplewindow_obj;

                // TS 23.02.18 dem principal.mailserverpassword eine kennung dazuschreiben dass es mit diesem kennwort funktioniert hat
                // und dadurch nicht mehr übertragen werden muss bis es irgenwann nicht geklappt hat denn dann wird die kennung wieder weggenommen                
                if (DataAdapter.Instance.DataCache.ResponseStatus.success)
                {
                    DataAdapter.Instance.DataCache.Rights.MarkUserPrincipalMailPasswordValidated();
                }

                if (_CanLC_MailCreate_IsLCAvail() && (DataAdapter.Instance.DataCache.ResponseStatus.success || !pushtolc))
                {
                    string lcexecutefile = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.LC_Command];
                    string lcparamfilename = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.LC_Param];
                    CSLCRequest lcrequest = CreateLCRequest();

                    // TS 14.12.15
                    string mailidcreated = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.MailId_Created];
                    lcrequest.proc_param_objectids = new string[1];
                    lcrequest.proc_param_objectids[0] = mailidcreated;
                    lcrequest.proc_param_showsampleswindow = showsamplewindow;

                    DataAdapter.Instance.DataProvider.LC_CreateCommandFile(CSEnumCommands.cmd_LC_MailShow, lcparamfilename, lcrequest, pushtolc, DataAdapter.Instance.DataCache.Rights.UserPrincipal);
                    // nur wenn nicht bereits mit push weitergeleitet wurde
                    if (!pushtolc && lcexecutefile != null && lcexecutefile.Length > 0 && lcparamfilename != null && lcparamfilename.Length > 0)
                    {
                        LC_ExecuteCall(lcexecutefile, lcparamfilename);
                    }
                }
            }
            catch (System.Exception ex) { if (ex.Message.Length > 0) { Log.Log.Debug(ex.Message); } }
        }

        private void LC_MailCreate_Done_SendClientCall(string mailserverclientcall)
        {
            // TS 23.02.18 dem principal.mailserverpassword eine kennung dazuschreiben dass es mit diesem kennwort funktioniert hat
            // und dadurch nicht mehr übertragen werden muss bis es irgenwann nicht geklappt hat denn dann wird die kennung wieder weggenommen                
            if (DataAdapter.Instance.DataCache.ResponseStatus.success)
            {
                DataAdapter.Instance.DataCache.Rights.MarkUserPrincipalMailPasswordValidated();
            }

            if (DataAdapter.Instance.DataCache.ResponseStatus.success && mailserverclientcall != null && mailserverclientcall.Length > 0)
            {
                // aufruf externe anwendung
                HtmlPage.Window.Navigate(new Uri(mailserverclientcall, UriKind.RelativeOrAbsolute), "_blank", "");
            }
        }

        #endregion LC_MailCreate (cmd_LC_MailCreate)

        // ==================================================================
        // mail beantworten/weiterleiten (mail wird über server in postfach eingestellt, lc zeigt es dann an)

        #region LC_MailShow (cmd_LC_MailShow)

        public DelegateCommand cmd_LC_MailShow { get { return new DelegateCommand(LC_MailShow, _CanLC_MailShow); } }

        public void LC_MailShow()
        {
            LC_MailShow(false, "");
        }

        // TS 30.09.15 einstieg über callback wg. login dialog
        private void LC_MailShow(List<object> parameters)
        {
            if (parameters.GetType().Name.StartsWith("List"))
            {
                List<object> paramlist = (List<object>)parameters;
                if (paramlist.Count >= 1)
                {
                    bool askcredentials = false;
                    if (paramlist[0] != null) askcredentials = (bool)paramlist[0];
                    string askedpassword = null;
                    if (paramlist.Count > 1 && paramlist[1] != null) askedpassword = (string)paramlist[1];

                    LC_MailShow(askcredentials, askedpassword);
                }
            }
        }

        public void LC_MailShow(bool askcredentials, string askedpassword)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                // TS 30.09.15
                List<object> callbackparams = new List<object>();

                if (askcredentials)
                {
                    // TS 30.09.15
                    if (DataAdapter.Instance.DataCache.ResponseStatus.returncode.Equals(Statics.Constants.SERVERERRORCODE_MAILCONNECT_FAIL)
                        ||
                        DataAdapter.Instance.DataCache.ResponseStatus.returncode.Equals(Statics.Constants.SERVERERRORCODE_MAILLISTENER_FAIL))
                    {
                        callbackparams.Add(false);
                        View.Dialogs.dlgAskPW askpw = new View.Dialogs.dlgAskPW(LC_MailShow, callbackparams, null);
                        DialogHandler.Show_Dialog(askpw);
                    }
                }
                else
                {
                    // TS 30.09.15 callback falls nicht an mailserver angemeldet werden konnte dann nochmal hier aufrufen mit askcredentials = true
                    callbackparams.Add(true);
                    CallbackAction callbackiffailed = new CallbackAction(LC_MailShow, callbackparams);

                    // TS 30.09.15 neuen principal verwenden wg. kennwortabfrage
                    // TS 11.01.17 principal erweitert, das mailserverpw ist nun als extra feld vorhanden
                    //CSUserPrincipal localprincipal = DataCache_Rights.CopyPrincipal(DataAdapter.Instance.DataCache.Rights.UserPrincipal);
                    //if (askedpassword != null && askedpassword.Length > 0) localprincipal.password = askedpassword;
                    CSUserPrincipal principal = DataAdapter.Instance.DataCache.Rights.UserPrincipal;
                    if (askedpassword != null && askedpassword.Length > 0) principal.mailserverpassword = askedpassword;

                    bool pushtolc = LC_IsPushEnabled;
                    string mailserverclientcall = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.LC_MailServerClientCall];
                    if (mailserverclientcall == null) mailserverclientcall = "";

                    IDocumentOrFolder currentdoc = DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected;
                    if (!currentdoc.isMailDocument)
                    {
                        currentdoc = null;
                        IDocumentOrFolder currentfld = DataAdapter.Instance.DataCache.Objects(Workspace).Folder_Selected;
                        foreach (IDocumentOrFolder child in currentfld.ChildDocuments)
                        {
                            if (child.isMailDocument)
                            {
                                currentdoc = child;
                                break;
                            }
                        }
                    }

                    // TS 18.05.16
                    if (currentdoc.isMailDocument)
                    {
                        // TS 30.09.15 zusammengefasst und mit failcallback
                        // wenn kein spezieller client call und nicht die push kommunikation verwendet wird dann gleich an lc weiterleiten
                        //if (mailserverclientcall.Length == 0 && !pushtolc)
                        //{
                        //    DataAdapter.Instance.DataProvider.RestoreMail(DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected.CMISObject, DataAdapter.Instance.DataCache.Rights.UserPrincipal);
                        //    LC_MailShow_Done_SendLC(pushtolc);
                        //}
                        //else
                        //{
                        //    // wenn kein spezieller client call aber die push kommunikation verwendet wird
                        //    if (mailserverclientcall.Length == 0)
                        //        DataAdapter.Instance.DataProvider.RestoreMail(DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected.CMISObject,
                        //            DataAdapter.Instance.DataCache.Rights.UserPrincipal, new CallbackAction(LC_MailShow_Done_SendLC, pushtolc));
                        //    else
                        //        DataAdapter.Instance.DataProvider.RestoreMail(DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected.CMISObject,
                        //            DataAdapter.Instance.DataCache.Rights.UserPrincipal, new CallbackAction(LC_MailShow_Done_SendClientCall, mailserverclientcall));
                        //}

                        // TS 14.12.15 alte mailid aus infocache vorher löschen
                        CSInformation[] infoarray = new CSInformation[1];
                        infoarray[0] = new CSInformation();
                        infoarray[0].informationid = CSEnumInformationId.MailId_Created;
                        infoarray[0].informationidSpecified = true;
                        infoarray[0].informationvalue = "";
                        DataAdapter.Instance.DataCache.ApplyData(infoarray);

                        if (mailserverclientcall.Length == 0)
                        {
                            // TS 18.05.16
                            //DataAdapter.Instance.DataProvider.RestoreMail(DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected.CMISObject,
                            //    localprincipal, new CallbackAction(LC_MailShow_Done_SendLC, pushtolc), callbackiffailed);
                            DataAdapter.Instance.DataProvider.RestoreMail(currentdoc.CMISObject, principal, new CallbackAction(LC_MailShow_Done_SendLC, pushtolc), callbackiffailed);
                        }
                        else
                        {
                            // TS 18.05.16
                            //DataAdapter.Instance.DataProvider.RestoreMail(DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected.CMISObject,
                            //    localprincipal, new CallbackAction(LC_MailShow_Done_SendClientCall, mailserverclientcall), callbackiffailed);
                            DataAdapter.Instance.DataProvider.RestoreMail(currentdoc.CMISObject, principal, new CallbackAction(LC_MailShow_Done_SendClientCall, mailserverclientcall), callbackiffailed);
                        }
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public bool CanLC_MailShow { get { return _CanLC_MailShow(); } }

        private bool _CanLC_MailShow()
        {
            bool ret = DataAdapter.Instance.DataCache.ApplicationFullyInit;
            if (ret)
            {
                ret = false;
                IDocumentOrFolder currentdoc = DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected;
                if (!currentdoc.isMailDocument)
                {
                    currentdoc = null;
                    IDocumentOrFolder currentfld = DataAdapter.Instance.DataCache.Objects(Workspace).Folder_Selected;
                    foreach (IDocumentOrFolder child in currentfld.ChildDocuments)
                    {
                        if (child.isMailDocument)
                        {
                            currentdoc = child;
                            break;
                        }
                    }
                }
                if (currentdoc != null)
                    ret = currentdoc.objectId.Length > 0 && currentdoc.isNotCreated == false;
            }
            return ret;
        }

        private bool _CanLC_MailShow_IsLCAvail()
        {
            return CanLC_MailShow && LC_IsAvailable && LC_IsModulAvailable(CSEnumInformationId.LC_Modul_Outlook, CSEnumInformationId.LC_Modul_OutlookParam);
        }

        private void LC_MailShow_Done_SendLC(bool pushtolc)
        {
            // TS 23.02.18 dem principal.mailserverpassword eine kennung dazuschreiben dass es mit diesem kennwort funktioniert hat
            // und dadurch nicht mehr übertragen werden muss bis es irgenwann nicht geklappt hat denn dann wird die kennung wieder weggenommen                
            if (DataAdapter.Instance.DataCache.ResponseStatus.success)
            {
                DataAdapter.Instance.DataCache.Rights.MarkUserPrincipalMailPasswordValidated();
            }

            if (_CanLC_MailShow_IsLCAvail() && (DataAdapter.Instance.DataCache.ResponseStatus.success || !pushtolc))
            {
                string lcexecutefile = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.LC_Command];
                string lcparamfilename = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.LC_Param];
                CSLCRequest lcrequest = CreateLCRequest();

                // TS 14.12.15
                string mailidcreated = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.MailId_Created];
                lcrequest.proc_param_objectids = new string[1];
                lcrequest.proc_param_objectids[0] = mailidcreated;

                DataAdapter.Instance.DataProvider.LC_CreateCommandFile(CSEnumCommands.cmd_LC_MailShow, lcparamfilename, lcrequest, pushtolc, DataAdapter.Instance.DataCache.Rights.UserPrincipal);
                // nur wenn nicht bereits mit push weitergeleitet wurde
                if (!pushtolc && lcexecutefile != null && lcexecutefile.Length > 0 && lcparamfilename != null && lcparamfilename.Length > 0)
                {
                    LC_ExecuteCall(lcexecutefile, lcparamfilename);
                }
            }
        }

        private void LC_MailShow_Done_SendClientCall(string mailserverclientcall)
        {
            // TS 23.02.18 dem principal.mailserverpassword eine kennung dazuschreiben dass es mit diesem kennwort funktioniert hat
            // und dadurch nicht mehr übertragen werden muss bis es irgenwann nicht geklappt hat denn dann wird die kennung wieder weggenommen                
            if (DataAdapter.Instance.DataCache.ResponseStatus.success)
            {
                DataAdapter.Instance.DataCache.Rights.MarkUserPrincipalMailPasswordValidated();
            }

            if (DataAdapter.Instance.DataCache.ResponseStatus.success && mailserverclientcall != null && mailserverclientcall.Length > 0)
            {
                // aufruf externe anwendung
                HtmlPage.Window.Navigate(new Uri(mailserverclientcall, UriKind.RelativeOrAbsolute), "_blank", "");
            }
        }

        #endregion LC_MailShow (cmd_LC_MailShow)

        // ==================================================================

        #region LC_PasteFromOutlook (cmd_LC_MailPasteFromOutlook)

        public DelegateCommand cmd_LC_MailPasteFromOutlook { get { return new DelegateCommand(LC_PasteFromOutlook, _CanLC_PasteFromOutlook); } }

        public void LC_PasteFromOutlook()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanLC_PasteFromOutlook)
                {
                    bool pushtolc = LC_IsPushEnabled;
                    string mailserverclientcall = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.LC_MailServerClientCall];
                    if (mailserverclientcall == null) mailserverclientcall = "";

                    string lcexecutefile = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.LC_Command];
                    string lcparamfilename = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.LC_Param];
                    CSLCRequest lcrequest = CreateLCRequest();

                    DataAdapter.Instance.DataProvider.LC_CreateCommandFile(CSEnumCommands.cmd_LC_MailPasteFromOutlook, lcparamfilename, lcrequest, pushtolc, DataAdapter.Instance.DataCache.Rights.UserPrincipal);
                    // nur wenn nicht bereits mit push weitergeleitet wurde
                    if (!pushtolc && lcexecutefile != null && lcexecutefile.Length > 0 && lcparamfilename != null && lcparamfilename.Length > 0)
                    {
                        LC_ExecuteCall(lcexecutefile, lcparamfilename);
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        public bool CanLC_PasteFromOutlook { get { return _CanLC_PasteFromOutlook(); } }

        private bool _CanLC_PasteFromOutlook()
        {
            return DataAdapter.Instance.DataCache.ApplicationFullyInit && LC_IsAvailable && LC_IsModulAvailable(CSEnumInformationId.LC_Modul_Outlook, CSEnumInformationId.LC_Modul_OutlookParam)
                && !DataAdapter.Instance.DataCache.Objects(this.Workspace).Folder_Selected.isNotCreated
                && DataAdapter.Instance.DataCache.Objects(this.Workspace).Folder_Selected.canCreateObjectLevel(Constants.STRUCTLEVEL_09_DOKLOG);
        }

        #endregion LC_PasteFromOutlook (cmd_LC_MailPasteFromOutlook)

        // ==================================================================

        #region LC_Scan (cmd_LC_Scan)

        //mr 01.07.2015 scan-change
        public DelegateCommand cmd_LC_Scan { get { return new DelegateCommand(LC_Scan, _CanLC_Scan); } }

        public void LC_Scan_Hotkey() { LC_Scan(); }
        public void LC_Scan()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanLC_Scan)
                {
                    string lcexecutefile = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.LC_Command];
                    string lcparamfilename = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.LC_Param];
                    CSLCRequest lcrequest = CreateLCRequest();
                    lcrequest.proc_param_multipage = true;

                    // TS 20.11.14
                    // lcrequest.proc_param_profilename = LC_ScanKofax_GetFirstProfileName();
                    string profilename = LC_Scan_CurrentProfile;
                    if (profilename == null || profilename.Length == 0)
                        profilename = LC_Scan_GetFirstProfileName();

                    lcrequest.proc_param_profilename = profilename;
                    lcrequest.proc_param_showdialog = 0;

                    // TS 27.11.14
                    string scanimagekey = "";
                    string reusableImageKey = "";
                    bool usescanimagekey = DataAdapter.Instance.DataCache.Profile.Option_GetValueBoolean(DataAdapter.Instance.DataCache.Profile.ProfileLayout.options, CSEnumOptions.autoconvert_scannedimages, true);
                    bool askForMergin = false;
                    if (usescanimagekey)
                    {
                        IDocumentOrFolder dummy = DataAdapter.Instance.DataCache.Objects(this.Workspace).Object_Selected;
                        if (dummy.isNotCreated && dummy.isDocument && dummy[CSEnumCmisProperties.scanImageKey] != null && dummy[CSEnumCmisProperties.scanImageKey].Length > 0)
                        {
                            // If the selected object is a new created document with a scanImageKey, then use it
                            reusableImageKey = dummy[CSEnumCmisProperties.scanImageKey];
                            askForMergin = true;
                        }
                        else
                        {
                            // If the selected object has child-documents, check for new documents with an existing scanImageKey
                            if (dummy.ChildDocuments.Count > 0)
                            {
                                foreach(IDocumentOrFolder childDoc in dummy.ChildDocuments)
                                {
                                    if(childDoc.isNotCreated && childDoc.isDocument && childDoc.ScanImageKey != null && childDoc.ScanImageKey.Length > 0)
                                    {
                                        reusableImageKey = childDoc.ScanImageKey;
                                        askForMergin = true;
                                        break;
                                    }
                                }
                            }
                            else
                            {
                                // Create a new Scanimage-Key
                                scanimagekey = System.DateTime.Now.ToString() + "_" + System.Environment.TickCount.ToString();
                            }
                        }

                        // Ask the user if a existing ScanimageKey should be used or a new will be created
                        if(askForMergin)
                        {
                            if(showYesNoDialog(LocalizationMapper.Instance["msg_lc_scan_merge_scanimagekeys"]))
                            {
                                scanimagekey = reusableImageKey;
                            }
                            else
                            {
                                // Create a new Scanimage-Key
                                scanimagekey = System.DateTime.Now.ToString() + "_" + System.Environment.TickCount.ToString();
                            }
                        }
                    }
                    lcrequest.proc_param_scanimagekey = scanimagekey;

                    // TS 05.12.14 zum scannen erstmal auf false gesetzt
                    lcrequest.def_autosave = false;

                    string useclickoncecalls = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.LC_UseClickOnceCalls];
                    if (useclickoncecalls == null) useclickoncecalls = "0";
                    bool pushtolc = !useclickoncecalls.Equals("1");

                    // Start a slideshow for the scan process
                    StartSlideShow();

                    DataAdapter.Instance.DataProvider.LC_CreateCommandFile(CSEnumCommands.cmd_LC_Scan, lcparamfilename, lcrequest, pushtolc, DataAdapter.Instance.DataCache.Rights.UserPrincipal);
                    // LC nur rufen nur wenn nicht bereits beim serveraufruf gepushed wurde
                    if (!pushtolc && lcexecutefile != null && lcexecutefile.Length > 0 && lcparamfilename != null && lcparamfilename.Length > 0)
                    {
                        LC_ExecuteCall(lcexecutefile, lcparamfilename);
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public bool CanLC_Scan { get { return _CanLC_Scan(); } }

        private bool _CanLC_Scan()
        {
            IDocumentOrFolder selObject = DataAdapter.Instance.DataCache.Objects(Workspace).Folder_Selected;
            return DataAdapter.Instance.DataCache.ApplicationFullyInit && LC_IsAvailable && LC_IsModulAvailable(CSEnumInformationId.LC_Modul_Scan, CSEnumInformationId.LC_Modul_ScanParam)
                && selObject.objectId != ""
                && selObject.objectId != DataAdapter.Instance.DataCache.Objects(Workspace).Root.objectId
                && selObject.structLevel >= Constants.STRUCTLEVEL_09_DOKLOG ? selObject.canCreateDocument : selObject.canCreateFolder;
        }

        #endregion LC_Scan (cmd_LC_Scan)

        // ==================================================================

        #region LC_ReScan (cmd_LC_ReScan)

        public DelegateCommand cmd_LC_ReScan { get { return new DelegateCommand(LC_ReScan, _CanLC_ReScan); } }

        public bool LC_Is_Rescan_InProgress { get { return _isLC_ReScan_InProgress; } }

        public string LC_Rescan_TargetId { get { return _LC_ReScan_TargetId; } }

        private bool _isLC_ReScan_InProgress = false;
        private string _LC_ReScan_TargetId = "";

        public void LC_ReScan()
        {
            _isLC_ReScan_InProgress = false;
            _LC_ReScan_TargetId = "";

            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanLC_ReScan)
                {
                    string lcexecutefile = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.LC_Command];
                    string lcparamfilename = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.LC_Param];
                    CSLCRequest lcrequest = CreateLCRequest();
                    lcrequest.proc_param_multipage = false; // This one is used to set the single-page-scan

                    // Getting Scan-Profile-Information
                    string profilename = LC_Scan_CurrentProfile;
                    if (profilename == null || profilename.Length == 0)
                        profilename = LC_Scan_GetFirstProfileName();
                    lcrequest.proc_param_profilename = profilename;
                    lcrequest.proc_param_showdialog = 0;

                    //Get Scan-Image-Key
                    string scanimagekey = "";
                    IDocumentOrFolder dummy = DataAdapter.Instance.DataCache.Objects(this.Workspace).Object_Selected;
                    if (dummy.isNotCreated && dummy.isDocument && dummy[CSEnumCmisProperties.scanImageKey] != null && dummy[CSEnumCmisProperties.scanImageKey].Length > 0)
                        scanimagekey = dummy[CSEnumCmisProperties.scanImageKey];
                    else
                    {
                        scanimagekey = dummy.RepositoryId + "_" + dummy.objectId;
                    }
                    lcrequest.proc_param_scanimagekey = scanimagekey;

                    // Setting up additional params
                    lcrequest.def_autosave = false;
                    string useclickoncecalls = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.LC_UseClickOnceCalls];
                    if (useclickoncecalls == null) useclickoncecalls = "0";
                    bool pushtolc = !useclickoncecalls.Equals("1");
                    lcrequest.proc_param_replaceobjectid = DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected.objectId;

                    // Create Command-File
                    DataAdapter.Instance.DataProvider.LC_CreateCommandFile(CSEnumCommands.cmd_LC_ReScan, lcparamfilename, lcrequest, pushtolc, DataAdapter.Instance.DataCache.Rights.UserPrincipal);
                    _isLC_ReScan_InProgress = true;
                    _LC_ReScan_TargetId = dummy.objectId;
                    Log.Log.WriteToBrowserConsole("Rescan started for objectId: " + _LC_ReScan_TargetId);

                    // Only call LC, if not pushed already earlier
                    if (!pushtolc && lcexecutefile != null && lcexecutefile.Length > 0 && lcparamfilename != null && lcparamfilename.Length > 0)
                    {
                        LC_ExecuteCall(lcexecutefile, lcparamfilename);
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public bool CanLC_ReScan { get { return _CanLC_ReScan(); } }

        private bool _CanLC_ReScan()
        {
            bool ret = false;
            bool isLCValid = DataAdapter.Instance.DataCache.ApplicationFullyInit && LC_IsAvailable && LC_IsModulAvailable(CSEnumInformationId.LC_Modul_Scan, CSEnumInformationId.LC_Modul_ScanParam);

            if (isLCValid)
            {
                bool isReScanValid = DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected != null
                    && DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected.objectId.Length > 0
                    && DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected.isInWork;
                ret = isReScanValid;
            }
            return ret;
        }

        public void LC_ReScan_AfterProcessing(IDocumentOrFolder scanPage)
        {
            ForceUpdateTreeviewForArgh();
            DataAdapter.Instance.InformObservers();

            // Cleaning up
            _isLC_ReScan_InProgress = false;
            _LC_ReScan_TargetId = "";
        }

        #endregion LC_ReScan (cmd_LC_ReScan)

        #region LC_ScanSettings (cmd_LC_ScanSettings)

        //mr 01.07.2015 scan-change
        public DelegateCommand cmd_LC_ScanSettings { get { return new DelegateCommand(LC_ScanSettings, _CanLC_ScanSettings); } }

        public void LC_ScanSettings()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanLC_ScanSettings)
                {
                    string lcexecutefile = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.LC_Command];
                    string lcparamfilename = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.LC_Param];
                    CSLCRequest lcrequest = CreateLCRequest();
                    lcrequest.proc_param_multipage = true;
                    lcrequest.proc_param_showdialog = 1;

                    // TS 06.02.15
                    //lcrequest.proc_param_profilename = "";

                    //mr 01.07.2015 scan-change
                    string profilename = LC_Scan_CurrentProfile;
                    if (profilename == null || profilename.Length == 0)
                        profilename = LC_Scan_GetFirstProfileName();

                    lcrequest.proc_param_profilename = profilename;

                    string useclickoncecalls = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.LC_UseClickOnceCalls];
                    if (useclickoncecalls == null) useclickoncecalls = "0";
                    bool pushtolc = !useclickoncecalls.Equals("1");

                    //mr 01.07.2015 scan-change
                    DataAdapter.Instance.DataProvider.LC_CreateCommandFile(CSEnumCommands.cmd_LC_ScanSettings, lcparamfilename, lcrequest, pushtolc, DataAdapter.Instance.DataCache.Rights.UserPrincipal);
                    // LC nur rufen nur wenn nicht bereits beim serveraufruf gepushed wurde
                    if (!pushtolc && lcexecutefile != null && lcexecutefile.Length > 0 && lcparamfilename != null && lcparamfilename.Length > 0)
                    {
                        LC_ExecuteCall(lcexecutefile, lcparamfilename);
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public bool CanLC_ScanSettings { get { return _CanLC_ScanSettings(); } }

        private bool _CanLC_ScanSettings()
        {
            return DataAdapter.Instance.DataCache.ApplicationFullyInit && LC_IsAvailable && LC_IsModulAvailable(CSEnumInformationId.LC_Modul_Scan, CSEnumInformationId.LC_Modul_ScanParam);
        }

        #endregion LC_ScanSettings (cmd_LC_ScanSettings)



        // ==================================================================
        #region LC_ForceStart (cmd_LC_ForceStart)

        public DelegateCommand cmd_LC_ForceStart { get { return new DelegateCommand(LC_ForceStart, _CanLC_ForceStart); } }

        public void LC_ForceStart()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanLC_ForceStart)
                {
                    Init.AppManager.AskStartLocalConnector(true);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public bool CanLC_ForceStart { get { return _CanLC_ForceStart(); } }

        private bool _CanLC_ForceStart()
        {
            return DataAdapter.Instance.DataCache.ApplicationFullyInit && isMasterApplication;
        }

        #endregion LC_ForceStart (cmd_LC_ForceStart)

        // ==================================================================

        #region LC_ScanSettingsExt (cmd_LC_ScanSettingsExt)

        public DelegateCommand cmd_LC_ScanSettingsExt { get { return new DelegateCommand(LC_ScanSettingsExt, _CanLC_ScanSettingsExt); } }

        public void LC_ScanSettingsExt()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanLC_ScanSettingsExt)
                {
                    string lcexecutefile = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.LC_Command];
                    string lcparamfilename = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.LC_Param];

                    CSLCRequest lcrequest = CreateLCRequest();
                    lcrequest.proc_param_multipage = true;
                    lcrequest.proc_param_showdialog = 1;

                    // TS 06.02.15
                    //lcrequest.proc_param_profilename = "";

                    //mr 01.07.2015 scan-change
                    string profilename = LC_Scan_CurrentProfile;
                    if (profilename == null || profilename.Length == 0)
                        profilename = LC_Scan_GetFirstProfileName();

                    lcrequest.proc_param_profilename = profilename;

                    string useclickoncecalls = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.LC_UseClickOnceCalls];
                    if (useclickoncecalls == null) useclickoncecalls = "0";
                    bool pushtolc = !useclickoncecalls.Equals("1");

                    DataAdapter.Instance.DataProvider.LC_CreateCommandFile(CSEnumCommands.cmd_LC_ScanSettingsExt, lcparamfilename, lcrequest, pushtolc, DataAdapter.Instance.DataCache.Rights.UserPrincipal);
                    // LC nur rufen nur wenn nicht bereits beim serveraufruf gepushed wurde
                    if (!pushtolc && lcexecutefile != null && lcexecutefile.Length > 0 && lcparamfilename != null && lcparamfilename.Length > 0)
                    {
                        LC_ExecuteCall(lcexecutefile, lcparamfilename);
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public bool CanLC_ScanSettingsExt { get { return _CanLC_ScanSettingsExt(); } }

        private bool _CanLC_ScanSettingsExt()
        {
            string lcextmenu = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.LC_ParamExt];
            bool enableextmenu = lcextmenu != null && lcextmenu.Equals("1");

            return DataAdapter.Instance.DataCache.ApplicationFullyInit && LC_IsAvailable && LC_IsModulAvailable(CSEnumInformationId.LC_Modul_Scan, CSEnumInformationId.LC_Modul_ScanParam) && enableextmenu;
        }

        #endregion LC_ScanSettingsExt (cmd_LC_ScanSettingsExt)

        #region LC_SetScanSource (cmd_LC_SetScanSource)

        //mr 01.07.2015 scan-change
        public DelegateCommand cmd_LC_SetScanSource { get { return new DelegateCommand(LC_SetScanSource, _CanLC_SetScanSource); } }

        public void LC_SetScanSource(object sourcename)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanLC_SetScanSource)
                {
                    if (sourcename != null)
                    {
                        LC_Scan_CurrentProfile = (string)sourcename;
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public bool CanLC_SetScanSource { get { return _CanLC_SetScanSource(); } }

        private bool _CanLC_SetScanSource()
        {
            return DataAdapter.Instance.DataCache.ApplicationFullyInit && LC_IsAvailable && LC_IsModulAvailable(CSEnumInformationId.LC_Modul_Scan, CSEnumInformationId.LC_Modul_ScanParam);
        }

        public string IsCheckedLC_SetScanSource { get { return LC_Scan_CurrentProfile; } }

        #endregion LC_SetScanSource (cmd_LC_SetScanSource)

        // ==================================================================

        #region LC_FileSelect (cmd_LC_FileSelect)

        public void LC_FileSelect(string filename)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanLC_FileSelect)
                {
                    string lcexecutefile = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.LC_Command];
                    string lcparamfilename = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.LC_Param];
                    CSLCRequest lcrequest = CreateLCRequest();

                    string useclickoncecalls = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.LC_UseClickOnceCalls];
                    if (useclickoncecalls == null) useclickoncecalls = "0";
                    bool pushtolc = !useclickoncecalls.Equals("1");

                    // Create Command-File
                    if (filename != null && filename.Length > 0)
                    {
                        string[] files = new string[1];

                        //files[0] = filename.Replace("\\","/");
                        files[0] = filename;

                        lcrequest.proc_param_filenames = files;
                        DataAdapter.Instance.DataProvider.LC_CreateCommandFile(CSEnumCommands.cmd_Link_CreateFSLink, lcparamfilename, lcrequest, pushtolc, DataAdapter.Instance.DataCache.Rights.UserPrincipal);
                    }
                    else
                    {
                        lcrequest.proc_param_showdialog = 1;
                        DataAdapter.Instance.DataProvider.LC_CreateCommandFile(CSEnumCommands.cmd_Link_CreateFSLink, lcparamfilename, lcrequest, pushtolc, DataAdapter.Instance.DataCache.Rights.UserPrincipal);
                    }

                    // Only call LC, if not pushed already earlier
                    if (!pushtolc && lcexecutefile != null && lcexecutefile.Length > 0 && lcparamfilename != null && lcparamfilename.Length > 0)
                    {
                        LC_ExecuteCall(lcexecutefile, lcparamfilename);
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        private bool CanLC_FileSelect { get { return _CanLC_FileSelect(); } }

        private bool _CanLC_FileSelect() { return DataAdapter.Instance.DataCache.ApplicationFullyInit && LC_IsAvailable; }

        #endregion LC_FileSelect



        #endregion LC_Init

        // sonstige nicht implementierte
        //cmd_LC_DocumentCreate
        private void LC_DocumentCreate(IDocumentOrFolder parent, IDocumentOrFolder mailaddressobj)
        {
            //Funktion BriefNeu
            // ggf. Adressdaten mitgeben
        }

        private void LC_DocumentCreateMany(IDocumentOrFolder parent)
        {
            //Serienbrief, Adressdaten als CSV mitgeben ?
        }

        //private void LC_MailCreate(List<IDocumentOrFolder> documents, IDocumentOrFolder mailaddressobj)
        //{
        //    //Funktion MailNeu (optional mit Datei(en))
        //}
        //private void LC_MailEdit(List<IDocumentOrFolder> documents)
        //{
        //    //Funktion MailNeu (optional mit Datei(en))
        //}
        private void LC_ScanTwain(IDocumentOrFolder parent)
        {
            //Funktion ScanTwain
            // ggf. zusammenfassen mit ScanKofax
        }

        private CSLCRequest CreateLCRequest()
        {
            CSLCRequest lcrequest = new CSLCRequest();
            // TS 16.12.15 default=false !!
            //lcrequest.def_autosave = true;
            lcrequest.def_autosave = false;

            // TS 28.02.14
            //lcrequest.def_defaultfolderid = DataAdapter.Instance.DataCache.Profile.Profile_GetInternalRootId(Constants.INTERNAL_DOCTYPE_ABLAGE);

            // TS 10.07.15 umgebogen auf stapel bis ablage vernünftig läuft
            //lcrequest.def_defaultfolderid = DataAdapter.Instance.DataCache.Profile.Profile_GetInternalRootId(CSEnumInternalObjectType.Ablage.ToString());
            //lcrequest.def_defaultrepositoryid = DataAdapter.Instance.DataCache.Repository(CSEnumProfileWorkspace.workspace_ablage).RepositoryInfo.repositoryId;

            // und weiter umgezogen auf dokumenteneingang (post)
            lcrequest.def_defaultfolderid = DataAdapter.Instance.DataCache.Profile.Profile_GetInternalRootId(CSEnumInternalObjectType.Stapel.ToString());
            lcrequest.def_defaultrepositoryid = DataAdapter.Instance.DataCache.Repository(CSEnumProfileWorkspace.workspace_stapel).RepositoryInfo.repositoryId;
            lcrequest.def_defaultworkspace = CSEnumProfileWorkspace.workspace_stapel;
            lcrequest.def_currentappid = DataAdapter.Instance.DataCache.Rights.ChoosenApplication.id;
            // TS 19.11.15
            lcrequest.def_defaultworkspaceSpecified = true;
            // und weiter umgezogen auf dokumenteneingang (post)
            List<string> subscriptions = DataAdapter.Instance.DataCache.Profile.Profile_GetSubscribedIds(Statics.Constants.POST_SUBSCRIBED_DEFAULT_COMMENT);
            if (subscriptions != null && subscriptions.Count > 0)
            {
                string parentid = subscriptions[0];
                string repositoryid = DataAdapter.Instance.DataCache.Repository(CSEnumProfileWorkspace.workspace_post).RepositoryInfo.repositoryId;
                lcrequest.def_defaultfolderid = parentid;
                lcrequest.def_defaultrepositoryid = repositoryid;
                lcrequest.def_defaultworkspace = CSEnumProfileWorkspace.workspace_post;
            }

            // TS 23.03.16 wenn clipboard dann weglassen weil parent root ist, sonst fehler
            //lcrequest.proc_repositoryid = DataAdapter.Instance.DataCache.Objects(this.Workspace).Folder_Selected.RepositoryId;
            //lcrequest.proc_parentobjectid = DataAdapter.Instance.DataCache.Objects(this.Workspace).Folder_Selected.objectId;
            IDocumentOrFolder current = DataAdapter.Instance.DataCache.Objects(this.Workspace).Folder_Selected;
            if (!current.Workspace.ToString().Equals(CSEnumProfileWorkspace.workspace_clipboard.ToString()))
            {
                lcrequest.proc_repositoryid = DataAdapter.Instance.DataCache.Objects(this.Workspace).Folder_Selected.RepositoryId;
                lcrequest.proc_parentobjectid = DataAdapter.Instance.DataCache.Objects(this.Workspace).Folder_Selected.objectId;
            }

            // TS 06.02.15
            lcrequest.proc_param_workspace = DataAdapter.Instance.DataCache.Objects(this.Workspace).Folder_Selected.Workspace;
            if (lcrequest.proc_param_workspace.ToString().Equals(CSEnumProfileWorkspace.workspace_undefined.ToString()))
                lcrequest.proc_param_workspace = this.Workspace;
            lcrequest.proc_param_workspaceSpecified = true;

            return lcrequest;
        }

        private void LC_ExecuteCall(string lcexecutefilename, string lcparamfilename)
        {
            // toolbar alle comboboxen schliessen
            LC_CloseAllToolbarCombos();

            //if (DataAdapter.Instance.DataCache.ResponseStatus.success)
            //{
            //string lcexecutefile = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.LC_Command];
            //string paramurl = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.LC_Param];
            if (lcparamfilename != null && lcparamfilename.Length > 0)
                lcparamfilename = Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(lcparamfilename));

            //MessageBox.Show(lcexecutefilename);

            // so gehts im test, aber nicht auf jboss
            //CustomHyperlinkButton ccc = new CustomHyperlinkButton();
            //ccc.Click += (sender, arg) =>
            //{
            //    System.Windows.Browser.HtmlPage.Window.Navigate(new Uri(lcexecutefilename + "?p=" + lcparamfilename), "_self");
            //    ccc = null;
            //};
            //ccc.OnClickPublic();

            CustomHyperlinkButton ccc = new CustomHyperlinkButton();
            ccc.NavigateUri = new Uri(lcexecutefilename + "?p=" + lcparamfilename);
            ccc.TargetName = "_self";
            ccc.OnClickPublic();
            //}
        }

        public bool LC_IsAvailable
        {
            get
            {
                string lcexecutefile = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.LC_Command];
                string lcparamfilename = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.LC_Param];
                return (lcexecutefile != null && lcexecutefile.Length > 0 && lcparamfilename != null && lcparamfilename.Length > 0);
            }
        }

        private bool LC_IsPushEnabled
        {
            get
            {
                string useclickoncecalls = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.LC_UseClickOnceCalls];
                if (useclickoncecalls == null) useclickoncecalls = "0";
                bool pushtolc = !useclickoncecalls.Equals("1");
                return pushtolc;
            }
        }

        public bool LC_IsModulAvailable(CSEnumInformationId lc_modul, CSEnumInformationId lc_modul_param)
        {
            string lcexecutefile = DataAdapter.Instance.DataCache.Info[lc_modul];
            string lcparamfilename = DataAdapter.Instance.DataCache.Info[lc_modul_param];
            return (lcexecutefile != null && lcexecutefile.Length > 0 && lcparamfilename != null && lcparamfilename.Length > 0);
        }

        public string LC_GetDocEditApplication(string docfilename)
        {
            string ret_application = _LC_GetDocEditApplication(docfilename, CSEnumInformationId.LC_Modul_Word, CSEnumInformationId.LC_Modul_WordParam);
            if (ret_application.Length == 0)
                ret_application = _LC_GetDocEditApplication(docfilename, CSEnumInformationId.LC_Modul_Excel, CSEnumInformationId.LC_Modul_ExcelParam);
            return ret_application;
        }

        private string _LC_GetDocEditApplication(string docfilename, CSEnumInformationId lc_modul, CSEnumInformationId lc_modul_param)
        {
            string ret_application = "";
            if (docfilename != null && docfilename.Length > 0)
            {
                string lcparam = DataAdapter.Instance.DataCache.Info[lc_modul_param];
                if (docfilename.Contains("."))
                {
                    string extension = docfilename.Substring(docfilename.LastIndexOf(".")).ToLower();

                    // TS 12.0o2.14 umgebaut auf standard parameter mit | getrennt
                    // if (lcparam != null && lcparam.Length > 0 && lcparam.Contains("(") && lcparam.Contains(")"))
                    if (lcparam != null && lcparam.Length > 0)
                    {
                        // TS 12.0o2.14 umgebaut auf standard parameter mit | getrennt
                        //int first = lcparam.IndexOf("(");
                        //int last = lcparam.IndexOf(")");
                        //lcparam = lcparam.Substring(first, last - first);
                        //char splitter = (",".ToCharArray())[0];
                        char splitter = ("|".ToCharArray())[0];
                        string[] lctokens = lcparam.Split(splitter);
                        foreach (string token in lctokens)
                        {
                            if (token.ToLower().Equals(extension))
                            {
                                ret_application = DataAdapter.Instance.DataCache.Info[lc_modul];
                                break;
                            }
                        }
                    }
                }
            }
            return ret_application;
        }

        //mr 01.07.2015 scan-change
        private string LC_Scan_GetFirstProfileName()
        {
            string ret = "";
            string lcparam = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.LC_Modul_ScanParam];
            if (lcparam != null && lcparam.Length > 0)
            {
                char splitter = ("|".ToCharArray())[0];
                string[] lctokens = lcparam.Split(splitter);
                if (lctokens.Length > 0)
                    ret = lctokens[0];
            }
            return ret;
        }

        public List<string> LC_Scan_GetProfileNames()
        {
            List<string> ret = new List<string>();
            string lcparam = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.LC_Modul_ScanParam];
            if (lcparam != null && lcparam.Length > 0)
            {
                char splitter = ("|".ToCharArray())[0];
                string[] lctokens = lcparam.Split(splitter);
                if (lctokens.Length > 0)
                {
                    foreach (string token in lctokens)
                        ret.Add(token);
                }
            }
            return ret;
        }

        public string LC_Scan_CurrentProfile
        {
            get
            {
                string ret = "";
                CSOption option = DataAdapter.Instance.DataCache.Profile.Option_GetOption_LevelUserProfile(CSEnumOptions.scansource);
                if (option != null)
                {
                    ret = (string)option.value;
                }
                return ret;
            }
            set
            {
                CSOption option = DataAdapter.Instance.DataCache.Profile.Option_GetOption_LevelUserProfile(CSEnumOptions.scansource);
                if (option != null)
                {
                    option.value = (string)value;
                    // TS 02.06.17 mitteilung an html client
                    ViewManager.SpreadScantemplate((string)value);
                }
            }
        }

        // damit die popup menus und button menus nicht offen bleiben alle durchrennen und schliessen
        private void LC_CloseAllToolbarCombos()
        {
            try
            {
                // toolbar alle comboboxen schliessen
                CSProfileComponent pc = DataAdapter.Instance.DataCache.Profile.Profile_GetFirstComponentOfType(CSEnumProfileComponentType.TOOLBAR);
                if (pc != null)
                {
                    IVisualObserver obs = DataAdapter.Instance.FindVisualObserverForProfileComponent(pc);
                    if (obs != null)
                    {
                        IPageBindingAdapter pba = obs.GetPageBindingAdapter();
                        if (pba != null && pba.PageProcessing != null)
                        {
                            WrapPanel commands = VisualTreeFinder.FindControlDown<WrapPanel>(pba.PageProcessing.Page.layoutRoot, typeof(WrapPanel), Statics.Constants.WRAPPANEL_COMMANDS, -1);
                            if (commands != null)
                            {
                                foreach (UIElement child in commands.Children)
                                {
                                    if (child.GetType().Name.Equals("ToggleButton"))
                                    {
                                        if (((System.Windows.Controls.Primitives.ToggleButton)child).Content != null)
                                        {
                                            // TS 22.07.15 typüberprüfung
                                            if (((System.Windows.Controls.Primitives.ToggleButton)child).Content.GetType().Name.Equals("ComboBox"))
                                            {
                                                ComboBox togglecontent = (ComboBox)((System.Windows.Controls.Primitives.ToggleButton)child).Content;
                                                if (togglecontent.IsDropDownOpen)
                                                {
                                                    togglecontent.IsDropDownOpen = false;
                                                    break;
                                                }
                                            }
                                        }
                                    }
                                    if (child.GetType().Name.Equals("ComboBox"))
                                    {
                                        if (((ComboBox)child).IsDropDownOpen)
                                        {
                                            ((ComboBox)child).IsDropDownOpen = false;
                                            break;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception) { }
        }

        #endregion Local Connector

        // ==================================================================

        #region LoadTheme

        public void LoadTheme(string theme)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                DataAdapter.Instance.DataProvider.GetTheme(theme, DataAdapter.Instance.DataCache.Rights.UserPrincipal, new CallbackAction(LoadTheme_Done));
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public void LoadTheme_Done()
        {
            ViewManager.LoadAllComponents(null);
        }

        #endregion LoadTheme

        // ==================================================================

        #region MarkObjectsAsRead (cmd_MarkObjectsAsRead)

        public DelegateCommand cmd_MarkObjectsAsRead { get { return new DelegateCommand(MarkObjectsAsRead, _CanMarkObjectsAsRead); } }

        public void MarkObjectsAsRead()
        {
            MarkObjectsAsRead(null);
        }

        public void MarkObjectsAsRead(List<IDocumentOrFolder> objects)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (_CanMarkObjectsAsRead(objects))
                {
                    string parentid = "";
                    List<cmisObjectType> cmisobjects = new List<cmisObjectType>();
                    if (objects == null)
                    {
                        // TS 30.07.13
                        objects = new List<IDocumentOrFolder>();
                        if (DataAdapter.Instance.DataCache.Objects(this.Workspace).Objects_Selected.Count > 0)
                        {
                            foreach (IDocumentOrFolder tmp in DataAdapter.Instance.DataCache.Objects(this.Workspace).Objects_Selected)
                                objects.Add(tmp);
                        }
                        else
                        {
                            IDocumentOrFolder selectedobject = DataAdapter.Instance.DataCache.Objects(this.Workspace).Folder_Selected;
                            objects.Add(selectedobject);
                        }
                    }

                    foreach (IDocumentOrFolder tmp in objects)
                    {
                        cmisobjects.Add(tmp.CMISObject);
                        tmp.MarkedAsReadOnSelection = true;
                        // aus autoliste löschen
                        if (DataAdapter.Instance.DataCache.ObjectToAutoSetReadByUser.Keys.Contains(tmp.objectId))
                            DataAdapter.Instance.DataCache.ObjectToAutoSetReadByUser.Remove(tmp.objectId);
                        // TS 12.05.17
                        parentid = tmp.parentId;
                    }
                    if (cmisobjects.Count > 0)
                    {
                        // TS 12.05.17
                        //DataAdapter.Instance.DataProvider.MarkObjectsAsRead(cmisobjects, true, DataAdapter.Instance.DataCache.Rights.UserPrincipal);
                        //CallbackAction cb = new CallbackAction(MarkObjectsAsRead_Done);
                        CallbackAction cb = new CallbackAction(MarkObjectsAsRead_Done, parentid);
                        DataAdapter.Instance.DataProvider.MarkObjectsAsRead(cmisobjects, true, DataAdapter.Instance.DataCache.Rights.UserPrincipal, cb);
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        private void MarkObjectsAsRead_Done(string objectid)
        {
            //DataAdapter.Instance.InformObservers(this.Workspace);
            DataAdapter.Instance.DataCache.Info.UpdateAlertNotification(objectid, this.Workspace, false);
        }

        public void MarkObjectsAsRead_PrepareAutomaticSet(string objectidtoautoset)
        {
            IDocumentOrFolder tmp = DataAdapter.Instance.DataCache.Objects(Workspace).GetObjectById(objectidtoautoset);
            MarkObjectsAsRead_PrepareAutomaticSet(tmp);
        }

        public void MarkObjectsAsRead_PrepareAutomaticSet(IDocumentOrFolder objecttoautoset)
        {
            if (objecttoautoset.isDocument)
            {
                // TS 19.12.13 MoveParent berücksichtigen
                //objecttoautoset = DataAdapter.Instance.DataCache.Objects(Workspace).GetObjectById(objecttoautoset.parentId);
                objecttoautoset = DataAdapter.Instance.DataCache.Objects(Workspace).GetObjectById(objecttoautoset.TempOrRealParentId);
            }

            if (objecttoautoset != null && objecttoautoset.objectId.Length > 0 && objecttoautoset.canSetReadFlag && !objecttoautoset.isReadByUser)
            {
                if (!DataAdapter.Instance.DataCache.ObjectToAutoSetReadByUser.ContainsKey(objecttoautoset.objectId))
                    DataAdapter.Instance.DataCache.ObjectToAutoSetReadByUser.Add(objecttoautoset.objectId, objecttoautoset);
            }
        }

        public void MarkObjectsAsRead_ProcessAutomaticSet(string objectidtoautoset)
        {
            IDocumentOrFolder tmp = DataAdapter.Instance.DataCache.Objects(Workspace).GetObjectById(objectidtoautoset);
            MarkObjectsAsRead_ProcessAutomaticSet(tmp);
        }

        private string _markobjectsasread_lastobjectid = "";
        public void MarkObjectsAsRead_ProcessAutomaticSet(IDocumentOrFolder objecttoautoset)
        {
            List<IDocumentOrFolder> objects = new List<IDocumentOrFolder>();

            if (objecttoautoset.isDocument)
            {
                // TS 19.12.13 MoveParent berücksichtigen
                //objecttoautoset = DataAdapter.Instance.DataCache.Objects(Workspace).GetObjectById(objecttoautoset.parentId);
                objecttoautoset = DataAdapter.Instance.DataCache.Objects(Workspace).GetObjectById(objecttoautoset.TempOrRealParentId);
            }

            if (DataAdapter.Instance.DataCache.ObjectToAutoSetReadByUser.Keys.Contains(objecttoautoset.objectId))
            {
                // This one is for preventing useless spamming the "markobjectasread"-function by clear-methods on adding new files to a document
                if (!_markobjectsasread_lastobjectid.Equals(objecttoautoset.objectId))
                {
                    objects.Add(objecttoautoset);
                    _markobjectsasread_lastobjectid = objecttoautoset.objectId;
                }
            }

            if (objects.Count > 0)
                MarkObjectsAsRead(objects);
        }

        // ------------------------------------------------------------------
        public bool CanMarkObjectsAsRead { get { return _CanMarkObjectsAsRead(); } }

        private bool _CanMarkObjectsAsRead()
        {
            return _CanMarkObjectsAsRead(null);
        }

        private bool _CanMarkObjectsAsRead(List<IDocumentOrFolder> objects)
        {
            bool canmark = DataAdapter.Instance.DataCache.ApplicationFullyInit;
            if (canmark && isMasterApplication) // Only allow, if this is the master-page
            {
                if (objects == null)
                {
                    objects = new List<IDocumentOrFolder>();
                    if (DataAdapter.Instance.DataCache.Objects(this.Workspace).Objects_Selected.Count > 0)
                    {
                        foreach (IDocumentOrFolder tmp in DataAdapter.Instance.DataCache.Objects(this.Workspace).Objects_Selected)
                            objects.Add(tmp);
                    }
                    else
                    {
                        IDocumentOrFolder selectedobject = DataAdapter.Instance.DataCache.Objects(this.Workspace).Folder_Selected;
                        objects.Add(selectedobject);
                    }
                }
                canmark = false;
                if (objects.Count > 0)
                {
                    foreach (IDocumentOrFolder tmp in objects)
                    {
                        if (tmp != null && tmp.canSetReadFlag && !tmp.isReadByUser)
                        {
                            canmark = true;
                            break;
                        }
                    }
                }
            }
            return canmark && isMasterApplication;
        }

        // ------------------------------------------------------------------

        #endregion MarkObjectsAsRead (cmd_MarkObjectsAsRead)

        // ==================================================================

        #region MarkObjectsAsUnread (cmd_MarkObjectsAsUnread)

        public DelegateCommand cmd_MarkObjectsAsUnread { get { return new DelegateCommand(MarkObjectsAsUnread, _CanMarkObjectsAsUnread); } }

        public void MarkObjectsAsUnread()
        {
            MarkObjectsAsUnread(null);
        }

        public void MarkObjectsAsUnread(List<IDocumentOrFolder> objects)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (_CanMarkObjectsAsUnread(objects))
                {
                    List<cmisObjectType> cmisobjects = new List<cmisObjectType>();
                    if (objects == null)
                    {
                        // TS 30.07.13
                        objects = new List<IDocumentOrFolder>();
                        if (DataAdapter.Instance.DataCache.Objects(this.Workspace).Objects_Selected.Count > 0)
                        {
                            foreach (IDocumentOrFolder tmp in DataAdapter.Instance.DataCache.Objects(this.Workspace).Objects_Selected)
                                objects.Add(tmp);
                        }
                        else
                        {
                            IDocumentOrFolder selectedobject = DataAdapter.Instance.DataCache.Objects(this.Workspace).Folder_Selected;
                            objects.Add(selectedobject);
                        }
                    }
                    foreach (IDocumentOrFolder tmp in objects)
                    {
                        cmisobjects.Add(tmp.CMISObject);
                        // aus autoliste löschen
                        if (DataAdapter.Instance.DataCache.ObjectToAutoSetReadByUser.Keys.Contains(tmp.objectId))
                            DataAdapter.Instance.DataCache.ObjectToAutoSetReadByUser.Remove(tmp.objectId);
                    }

                    if (cmisobjects.Count > 0)
                        DataAdapter.Instance.DataProvider.MarkObjectsAsRead(cmisobjects, false, DataAdapter.Instance.DataCache.Rights.UserPrincipal);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        public bool CanMarkObjectsAsUnread { get { return _CanMarkObjectsAsUnread(); } }

        private bool _CanMarkObjectsAsUnread()
        {
            return _CanMarkObjectsAsUnread(null);
        }

        private bool _CanMarkObjectsAsUnread(List<IDocumentOrFolder> objects)
        {
            bool canmark = DataAdapter.Instance.DataCache.ApplicationFullyInit;
            if (canmark)
            {
                if (objects == null)
                {
                    objects = new List<IDocumentOrFolder>();
                    if (DataAdapter.Instance.DataCache.Objects(this.Workspace).Objects_Selected.Count > 0)
                    {
                        foreach (IDocumentOrFolder tmp in DataAdapter.Instance.DataCache.Objects(this.Workspace).Objects_Selected)
                            objects.Add(tmp);
                    }
                    else
                    {
                        IDocumentOrFolder selectedobject = DataAdapter.Instance.DataCache.Objects(this.Workspace).Folder_Selected;
                        objects.Add(selectedobject);
                    }
                }
                canmark = false;
                if (objects.Count > 0)
                {
                    foreach (IDocumentOrFolder tmp in objects)
                    {
                        if (tmp != null && tmp.canSetReadFlag && tmp.isReadByUser)
                        {
                            canmark = true;
                            break;
                        }
                    }
                }
            }
            return canmark;
        }

        // ------------------------------------------------------------------

        #endregion MarkObjectsAsUnread (cmd_MarkObjectsAsUnread)

        // ==================================================================

        #region MFPDelete (cmd_MFPDelete)

        public DelegateCommand cmd_MFPDelete { get { return new DelegateCommand(MFPDelete, _CanMFPDelete); } }

        public void MFPDelete()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanMFPDelete)
                {
                    // TS 18.01.16
                    // if (showYesNoDialog(Localization.localstring.msg_RequestForDelete))
                    if (showYesNoDialog(LocalizationMapper.Instance["msg_RequestForDelete"]))
                    {
                        List<cmisObjectType> todelete = new List<cmisObjectType>();
                        foreach (IDocumentOrFolder obj in DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_clipboard).Objects_Selected)
                            todelete.Add(obj.CMISObject);
                        DataAdapter.Instance.DataProvider.MFPDelete(todelete, DataAdapter.Instance.DataCache.Rights.UserPrincipal, new CallbackAction(MFPDelete_Done));
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        private void MFPDelete_Done()
        {
            // TS 15.02.16
            if (DataAdapter.Instance.DataCache.ResponseStatus.success)
                ClearClipboardSelected();
        }

        // ------------------------------------------------------------------
        public bool CanMFPDelete { get { return _CanMFPDelete(); } }

        private bool _CanMFPDelete()
        {
            bool ret = false;

            if (DataAdapter.Instance.DataCache.ApplicationFullyInit)
            {
                bool anymfp = false;
                bool anynotmfp = false;

                foreach (IDocumentOrFolder obj in DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_clipboard).Objects_Selected)
                {
                    if (obj.isMFPFile)
                        anymfp = true;
                    else
                        anynotmfp = true;
                }
                ret = (anymfp == true && anynotmfp == false);
            }
            return ret;
        }

        #endregion MFPDelete (cmd_MFPDelete)

        // ==================================================================

        #region MFPImport (cmd_MFPImport)

        public DelegateCommand cmd_MFPImport { get { return new DelegateCommand(MFPImport, _CanMFPImport); } }

        public void MFPImport()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanMFPImport)
                {
                    string customPath = DataAdapter.Instance.DataCache.Profile.Option_GetValue(CSEnumOptions.custommfppath);

                    // das root object von workspace_clipboard hat repository_id=0
                    // daher das root object vom default workspace nehmen
                    DataAdapter.Instance.DataProvider.MFPImport(DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_default).Root.CMISObject,
                        CSEnumProfileWorkspace.workspace_clipboard, customPath, DataAdapter.Instance.DataCache.Rights.UserPrincipal);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        public bool CanMFPImport { get { return _CanMFPImport(); } }

        private bool _CanMFPImport()
        {
            bool ret = false;
            if(DataAdapter.Instance.DataCache.ApplicationFullyInit)
            {
                IDocumentOrFolder selObj = DataAdapter.Instance.DataCache.Objects(this.Workspace).Folder_Selected;
                if (selObj.ACL != null)
                {
                    ret = selObj.hasCreatePermission(); //!selObj.hasReadOnlyPermission();
                }
                ret = ret && selObj.structLevel >= Constants.STRUCTLEVEL_09_DOKLOG ? selObj.canCreateDocument : selObj.canCreateFolder;
            }

            return ret;
        }

        #endregion MFPImport (cmd_MFPImport)

        // ==================================================================

        #region Move

        public DelegateCommand cmd_MoveDMS { get { return new DelegateCommand(MoveDMS, _CanMoveDMS); } }

        public void MoveDMS()
        {
            MoveDMS(false);
        }
        public void MoveDMS(bool autosave)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanMoveDMS)
                {
                    // If working on the master-applikation, get the selected default-workspace-object
                    IDocumentOrFolder target = null;
                    if (isMasterApplication)
                    {
                        target = DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_default).Folder_Selected;
                    }
                    else
                    // Get the default-workspace-object with the gathered ID from Communicator
                    {
                        string targetIDMaster = DataAdapter.Instance.DataCache.ClientCom.Last_MasterSelection_Default;
                        if (targetIDMaster != null && targetIDMaster.Length > 0)
                        {
                            target = DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_default).GetObjectById(targetIDMaster);
                        }
                    }

                    // And send the actual move-order
                    List<IDocumentOrFolder> source = DataAdapter.Instance.DataCache.Objects(this.Workspace).Objects_Selected;
                    Log.Log.WriteToBrowserConsole("MoveDMS for Source: " + source[0].objectId + " into target: " + target.objectId);

                    CallbackAction callback = null;
                    if (!isMasterApplication || autosave)
                    {
                        callback = new CallbackAction(DataAdapter.Instance.Processing(CSEnumProfileWorkspace.workspace_default).Save, new CallbackAction(DataAdapter.Instance.Processing(CSEnumProfileWorkspace.workspace_default).MoveDMS_Done_UpdateMaster, source));
                    }else
                    {
                        callback = new CallbackAction(MoveDMS_Done,null);
                    }

                    Move(source, target, false, callback);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        public bool CanMoveDMS { get { return _CanMoveDMS(); } }

        private bool _CanMoveDMS()
        {
            //return DataAdapter.Instance.DataCache.ApplicationFullyInit;
            bool ret = DataAdapter.Instance.DataCache.ApplicationFullyInit;

            if (ret)
            {
                // If working on the master-applikation, get the selected default-workspace-object
                IDocumentOrFolder target = null;
                if (isMasterApplication)
                {
                    target = DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_default).Folder_Selected;
                }
                else
                // Get the default-workspace-object with the gathered ID from Communicator
                {
                    string targetIDMaster = DataAdapter.Instance.DataCache.ClientCom.Last_MasterSelection_Default;
                    if (targetIDMaster != null && targetIDMaster.Length > 0)
                    {
                        target = DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_default).GetObjectById(targetIDMaster);
                    }
                }

                // TS 18.03.16
                ret = ret && target != null && !target.isNotCreated;

                foreach (IDocumentOrFolder obj in DataAdapter.Instance.DataCache.Objects(this.Workspace).Objects_Selected)
                {
                    if (ret) ret = obj.canMoveObject && target.canCreateObjectLevel(obj.structLevel);
                }
            }
            return ret;
        }

        private void MoveDMS_Done(CallbackAction finalCallback)
        {
            if (DataAdapter.Instance.DataCache.ResponseStatus.success)
            {
                DataAdapter.Instance.DataCache.Info.ListView_ForcePageSelection = true;
                DataAdapter.Instance.Processing(CSEnumProfileWorkspace.workspace_default).SetSelectedObject(DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_default).ObjectIdList_LastAddedFolders.Last());
                if (finalCallback != null)
                {
                    finalCallback.Invoke();
                }
            }
        }

        public void MoveDMS_Done_UpdateMaster(object sourcelist)
        {
            List<IDocumentOrFolder> source_list = (List<IDocumentOrFolder>)sourcelist;
            foreach (IDocumentOrFolder source in source_list)
            {
                ClientCommunication.Communicator.Instance.Send_UpdateMasterSlave(source.objectId);
            }
        }

        /// <summary>
        /// wird verwendet nur aus Drag & Drop
        /// es wird sofort gespeichert (je nach Optionseinstellung)
        /// </summary>
        /// <param name="objecttomove"></param>
        /// <param name="targetobject"></param>
        //public void Move(List<IDocumentOrFolder> objectstomove, IDocumentOrFolder targetobject)
        //{
        //    bool autosave = DataAdapter.Instance.DataCache.Profile.Option_GetValueBoolean(CSEnumOptions.autosave_drop);
        //    Move(objectstomove, targetobject, autosave);
        //}
        public void Move(List<IDocumentOrFolder> objectstomove, IDocumentOrFolder targetobject) { Move(objectstomove, targetobject, true, null); }

        public void Move(List<IDocumentOrFolder> objectstomove, IDocumentOrFolder targetobject, bool refreshtarget, CallbackAction finalCallback)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanMove(objectstomove, targetobject))
                {
                    List<cmisObjectType> cmisobjects = new List<cmisObjectType>();
                    bool isdocuments = false;
                    IDocumentOrFolder oldparent = null;
                    IDocumentOrFolder finaldoc = null;
                    CSEnumProfileWorkspace oldparentws = CSEnumProfileWorkspace.workspace_default;
                    CSEnumProfileWorkspace newparentws = targetobject.Workspace;
                    foreach (IDocumentOrFolder obj in objectstomove)
                    {
                        // TS 12.11.13 auf original umbiegen falls dok und nicht original
                        //cmisobjects.Add(obj.CMISObject);
                        IDocumentOrFolder objtest = GetValidCutObject(obj);

                        oldparentws = objtest.Workspace;
                        cmisobjects.Add(objtest.CMISObject);
                        // TS 19.12.13 hier soll ParentId bewusst so bleiben, prüfen ?!
                        oldparent = DataAdapter.Instance.DataCache.Objects(oldparentws).GetObjectById(objtest.parentId);
                        if (objtest.isDocument)
                        {
                            isdocuments = true;
                            finaldoc = objtest;
                        }
                    }

                    cmisObjectType cmistarget = null;
                    bool autosave = DataAdapter.Instance.DataCache.Profile.Option_GetValueBoolean(CSEnumOptions.autosave_drop);
                    if (targetobject != null)
                    {
                        cmistarget = targetobject.CMISObject;
                        if(objectstomove[objectstomove.Count - 1].Workspace != this.Workspace)
                        {
                            // TS 14.11.13 mit callback zum aktualisieren, aber nur wenn es dokumente sind  
                            if (isdocuments)
                            {
                                CallbackAction newparentcallfld = new CallbackAction(DataAdapter.Instance.Processing(targetobject.Workspace).SetSelectedObject, targetobject.objectId, 0, finalCallback);
                                CallbackAction newparentcall = new CallbackAction(DataAdapter.Instance.Processing(targetobject.Workspace).SetSelectedObject, finaldoc.objectId, 0, newparentcallfld);
                                // wenn sich die workspaces unterscheiden dann beide informieren, sonst nur den target  
                                if (!oldparentws.ToString().Equals(newparentws.ToString()))
                                {
                                    CallbackAction parentcalls = new CallbackAction(DataAdapter.Instance.Processing(oldparentws).SetSelectedObject, oldparent.objectId, 0, newparentcall);
                                    // TS 30.09.15 wenn keine aktualisierung gewünscht ist dann weglassen (z.b. bei MoveDMS)  
                                    if (!refreshtarget) parentcalls = null;
                                    DataAdapter.Instance.DataProvider.Move(cmisobjects, cmistarget, autosave, DataAdapter.Instance.DataCache.Rights.UserPrincipal, parentcalls);
                                }
                                else
                                {
                                    // TS 30.09.15 wenn keine aktualisierung gewünscht ist dann weglassen (z.b. bei MoveDMS)  
                                    if (!refreshtarget) newparentcall = null;
                                    DataAdapter.Instance.DataProvider.Move(cmisobjects, cmistarget, autosave, DataAdapter.Instance.DataCache.Rights.UserPrincipal, newparentcall);
                                }
                            }
                            else
                            {
                                // TS 02.07.15 auch dem verschobenen folgen  
                                string objectid = objectstomove[objectstomove.Count - 1].objectId;
                                CallbackAction callback = new CallbackAction(DataAdapter.Instance.Processing(targetobject.Workspace).SetSelectedObject, objectid, 0, finalCallback);

                                // TS 30.09.15 wenn keine aktualisierung gewünscht ist dann weglassen (z.b. bei MoveDMS)  
                                if (!refreshtarget) callback = null;
                                DataAdapter.Instance.DataProvider.Move(cmisobjects, cmistarget, autosave, DataAdapter.Instance.DataCache.Rights.UserPrincipal, callback);
                            }
                        }
                        else
                        {
                            DataAdapter.Instance.DataCache.Objects(this.Workspace).SetSelectedObject("");
                            CallbackAction callback = isMasterApplication == true ? finalCallback : new CallbackAction(MoveDMS_Done, finalCallback);
                            DataAdapter.Instance.DataProvider.Move(cmisobjects, cmistarget, autosave, DataAdapter.Instance.DataCache.Rights.UserPrincipal, callback);
                        }
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        public bool CanMove(List<IDocumentOrFolder> objectstomove, IDocumentOrFolder targetobject)
        {
            bool canmove = DataAdapter.Instance.DataCache.ApplicationFullyInit;
            // TS 18.03.16
            canmove = canmove && !targetobject.isNotCreated;

            foreach (IDocumentOrFolder obj in objectstomove)
            {
                if (canmove == true) canmove = targetobject.canCreateObjectLevel(obj.structLevel);

                // Check for unchanged-parent
                if (canmove == true) canmove = !targetobject.objectId.Equals(obj.parentId);

                // Additional check for pending jobs of some types
                if (canmove == true) canmove = !WSPendingJobHelper.ExistsCriticalJobs(obj.objectId);
            }
            if (canmove == true) canmove = _CanCut(objectstomove);

            return canmove;
        }

        /// <summary>
        /// info an alle
        /// </summary>
        //private void Move_Done(IDocumentOrFolder newparent, IDocumentOrFolder oldparent)
        //{
        //    if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();

        //    //if (oldparent.i

        //    DataAdapter.Instance.InformObservers();
        //    if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        //}
        // ------------------------------------------------------------------

        #endregion Move

        // ==================================================================

        #region SetNewOrderInParent                   


        public void SetNewOrderInParent(object srcObject, object targetObject)
        {
            try
            {
                // On Folders use the SetCustomSortOrder-Webservice Method
                if (!((IDocumentOrFolder)srcObject).isDocument)
                {
                    // Build orderby-string for reordering
                    string orderby = "";
                    List<CSOrderToken> sortTokens = TreeOrderTokens;
                    foreach (CServer.CSOrderToken orderbytoken in sortTokens)
                    {
                        if (orderby.Length > 0)
                        {
                            orderby = orderby + ", ";
                        }
                        orderby = orderby + orderbytoken.propertyname + " " + orderbytoken.orderby;
                    }

                    // Make sure, the order is given in sortpriority 1!
                    if(!orderby.ToUpper().Contains("FOLGE"))
                    {
                        orderby = "FOLGE asc, " + orderby;
                    }
                    DataAdapter.Instance.DataProvider.SetCustomSortOrder((IDocumentOrFolder)srcObject, (IDocumentOrFolder)targetObject, orderby, DataAdapter.Instance.DataCache.Rights.UserPrincipal, new CallbackAction(SetNewOrderInParent_Done, srcObject));
                }else
                // On Documents, change the FOLGE-Property manually and send it via UpdateProperties
                {
                    IDocumentOrFolder target = (IDocumentOrFolder)targetObject;
                    IDocumentOrFolder source = (IDocumentOrFolder)srcObject;
                    List<IDocumentOrFolder> saveList = new List<IDocumentOrFolder>();
                    string tmpnr = "";
                    string technr = "";
                    string inhaltnr = "";
                    string orinr = "";
                    int targetOrder = GetParsedStringFolge(target.Folge);
                    if (target.Folge.Length >= 7)
                    {
                        tmpnr = target.Folge.Substring(target.Folge.Length - 1, 1);
                        technr = target.Folge.Substring(target.Folge.Length - 3, 2);
                        inhaltnr = target.Folge.Substring(target.Folge.Length - 6, 3);
                        orinr = target.Folge.Substring(0, target.Folge.Length - 6);
                        targetOrder = GetParsedStringFolge(orinr);
                    }
                                        
                    // Create format string
                    string format = "";
                    for (int counter = 0; counter < orinr.Length; counter++)
                    {
                        format += "0";
                    }

                    // Attach the new order to the source-object marked with a "-" as trigger for the requester
                    if (targetOrder != 0)
                    {
                        string newfullorder = targetOrder.ToString(format) + inhaltnr + technr + tmpnr;
                        source.Folge = "-" + newfullorder;
                        source[CSEnumCmisProperties.isEdited] = "true";
                        saveList.Add(source);
                    }
                    if(saveList.Count > 0)
                    {
                        Save(saveList, new CallbackAction(SetNewOrderInParent_Done, srcObject));
                    }
                }
            }
            catch (System.Exception ex) { if (ex.Message.Length > 0) { Log.Log.Debug(ex.Message); } }          
        }

        private int GetParsedStringFolge(string toParse)
        {
            return int.Parse(toParse == null ? "1" : toParse.Length == 0 ? "1" : toParse);
        }


        private void SetNewOrderInParent_Done(object originalSource)
        {
            IDocumentOrFolder origSrc = (IDocumentOrFolder)originalSource;
            if(origSrc != null && origSrc.objectId.Length > 0)
            {
                DataAdapter.Instance.DataProvider.GetObjects(origSrc.Parent.ChildObjects.ToList(), DataAdapter.Instance.DataCache.Rights.UserPrincipal, new CallbackAction(ForceUpdateTreeviewForArgh));
            }
        }

        public void ForceUpdateTreeviewForArgh()
        {
            ForceUpdateTreeviewForArgh(null);
        }

        public void ForceUpdateTreeviewForArgh(CallbackAction callback)
        {
            Deployment.Current.Dispatcher.BeginInvoke((ThreadStart)(() =>
            {
                DataAdapter.Instance.DataCache.Info.RefreshTreeSortOrder = true;
                DataAdapter.Instance.InformObservers();
                if(callback != null)
                {
                    callback.Invoke();
                }
            }));
        }
        #endregion

        // ==================================================================
        #region Process_RC_QueryMandant

        /// <summary>
        /// Query-Mandant needs a list of query Key-Value-Pairs.
        /// Each KeyValue-Pair is separated with '#|#'
        /// Between a Key and a Value is a '#=#'
        /// One of the Query-Items may be the target mandant-id with the Key "MANDANT_ID"
        /// Example: 'MANDANT_ID#=#1#|#AKTE_01#=#Test#|#AKTESYS_07#=#Bobo'
        /// </summary>
        /// <param name="oqueryvalues"></param>
        /// <param name="oworkspace"></param>
        public void Process_RC_QueryMandant(object oqueryvalues, object oworkspace)
        {
            List<string> queryValues = (List<string>)oqueryvalues;
            CSEnumProfileWorkspace workspace = (CSEnumProfileWorkspace)oworkspace;
            string[] tokenSplitter = new string[] { "#|#" };
            string[] propertySplitter = new string[] { "#=#" };
            string mandantID = "";
            Dictionary<string, string> paramList = new Dictionary<string, string>();

            // Fetch Query-Tokens
            if (queryValues.Count > 0)
            {
                // Only one list of properties is supported
                string queryString = queryValues[0];
                string[] tokens = queryString.Split(tokenSplitter, StringSplitOptions.None);

                // Loop the splitted property-entries and build the query-object
                foreach (string querytoken in tokens)
                {
                    // 0 = Property-Name; 1 = Property-Value
                    string[] propertySet = querytoken.Split(propertySplitter, StringSplitOptions.None);
                    string propName = propertySet[0];
                    string propVal = propertySet[1];

                    if (propName.ToUpper().Equals("MANDANT_ID"))
                    {
                        // Read out the mandant-id
                        mandantID = propVal;
                    }
                    else
                    {
                        paramList.Add(propName, propVal);
                    }
                }
            }

            // Check for a necessary mandant-change
            if (mandantID.Length > 0 && !DataAdapter.Instance.DataCache.Rights.UserPrincipal.mandantid.Equals(mandantID))
            {
                // Create the new mandant, change the mandant, and start a new query, after all is done
                CSRightsMandant newMandant = new CSRightsMandant();
                foreach (CSRightsMandant m in DataAdapter.Instance.DataCache.Rights.UserRights.mandanten)
                {
                    if (m.id.Equals(mandantID))
                    {
                        newMandant = m;
                        break;
                    }
                }

                // Create the final callback and trigger the reload of the application
                CallbackAction cb = new CallbackAction(Process_RC_QueryMandant_Done, workspace, paramList);
                AppManager.StartupFinalCallback = cb;
                AppManager.ReloadAll(newMandant);
            }
            else
            {
                // Start the query directly
                Process_RC_QueryMandant_Done(workspace, paramList);
            }
        }

        private void Process_RC_QueryMandant_Done(object workspace, object query_params)
        {
            Dictionary<string, string> queryProperties = (Dictionary<string, string>)query_params;
            CSEnumProfileWorkspace castWorkspace = (CSEnumProfileWorkspace)workspace;

            // Apply the Search-Properties
            // Create the Property for the given index (name- & value-pair)
            foreach (KeyValuePair<string, string> kv in queryProperties)
            {
                cmisTypeContainer container = DataAdapter.Instance.DataCache.Repository(castWorkspace).GetTypeContainerForPropertyName(kv.Key);
                if (container != null && container.type.id != null && container.type.id.Length > 0)
                {
                    int structlevel = StructLevelFinder.GetStructLevelFromTypeId(container.type.id);
                    for (int i = 0; i < DataAdapter.Instance.DataCache.Objects(castWorkspace).ObjectList_EmptyQueryObjects.Count; i++)
                    {
                        IDocumentOrFolder emptyqueryobject = DataAdapter.Instance.DataCache.Objects(castWorkspace).ObjectList_EmptyQueryObjects[i];
                        if (emptyqueryobject.structLevel == structlevel)
                        {
                            emptyqueryobject[kv.Key] = kv.Value;
                            break;
                        }
                    }
                }
            }

            // Start the query
            DataAdapter.Instance.Processing(castWorkspace).Query();
        }

        #endregion

        // ==================================================================

        #region Paste (cmd_Paste)

        public DelegateCommand cmd_Paste { get { return new DelegateCommand(Paste, _CanPaste); } }

        // TS umbau auf mehrere
        //private void Paste() { Paste(DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_clipboard).Object_Selected, null); }
        private void Paste() { Paste(DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_clipboard).Objects_Selected, null); }

        /// <summary>
        /// CServer unterscheidet zwischen Copy oder Move und ruft diese
        /// Kein automatisches Speichern
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="target"></param>
        // TS 22.04.15 umbau auf mehrere
        //public void Paste(IDocumentOrFolder obj, IDocumentOrFolder target)
        public void Paste(List<IDocumentOrFolder> objectlist, IDocumentOrFolder target)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanPaste)
                {
                    // TS 01.04.15 wenn paste im bereich clipboard aufgerufen wird dann weiterleiten
                    // TS 20.04.16 bugfix: wenn bereits im workspace clipboard dann nicht weiterleiten sonst stackoverflow
                    // if (Workspace == CSEnumProfileWorkspace.workspace_clipboard)
                    if (Workspace == CSEnumProfileWorkspace.workspace_clipboard && !LastSelectedWorkspace.ToString().Equals(CSEnumProfileWorkspace.workspace_clipboard.ToString()))
                    {
                        // TS 22.04.15 umbau auf mehrere
                        //DataAdapter.Instance.Processing(LastSelectedWorkspace).Paste(DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_clipboard).Object_Selected, null);
                        DataAdapter.Instance.Processing(LastSelectedWorkspace).Paste(DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_clipboard).Objects_Selected, null);
                    }
                    else
                    {
                        List<cmisObjectType> pastecmisobjects = new List<cmisObjectType>();
                        // TS 22.04.15 umbau auf mehrere
                        foreach (IDocumentOrFolder singleobj in objectlist)
                        {
                            pastecmisobjects.Add(singleobj.CMISObject);
                        }                        

                        // Special behaviour for copy & paste into rootnodes
                        if (LastSelectedWorkspace != CSEnumProfileWorkspace.workspace_aktenplan || LastSelectedRootNode == null)
                        {
                            bool autosave = DataAdapter.Instance.DataCache.Profile.Option_GetValueBoolean(CSEnumOptions.autosave_paste);
                            if (target == null)
                            {
                                target = DataAdapter.Instance.DataCache.Objects(Workspace).Folder_Selected;
                            }
                            if (target == null || target.objectId.Length == 0)
                            {
                                target = DataAdapter.Instance.DataCache.Objects(Workspace).Root;
                            }
                            cmisObjectType cmistarget = target.CMISObject;

                            // TS 22.04.15
                            IDocumentOrFolder obj = objectlist[objectlist.Count - 1];

                            // TS 08.02.16 zurückgebaut
                            CallbackAction callback = new CallbackAction(DataAdapter.Instance.Processing(target.Workspace).ObjectTransfer_Done, obj.isDocument);
                            if (!obj.isDocument)
                            {
                                DataAdapter.Instance.DataProvider.Paste(pastecmisobjects, cmistarget, autosave, DataAdapter.Instance.DataCache.Rights.UserPrincipal, callback);
                            }
                            else
                            {
                                if (target.structLevel == Constants.STRUCTLEVEL_09_DOKLOG)
                                {
                                    DataAdapter.Instance.DataProvider.Paste(pastecmisobjects, cmistarget, autosave, DataAdapter.Instance.DataCache.Rights.UserPrincipal, callback);
                                }
                                else
                                {
                                    cmisTypeContainer foldertypedefinition = DataAdapter.Instance.DataCache.Repository(Workspace).GetTypeContainerForStructLevel(Constants.STRUCTLEVEL_09_DOKLOG, target);
                                    DataAdapter.Instance.DataProvider.PasteDocumentsInFolders(foldertypedefinition, cmistarget, pastecmisobjects, autosave, DataAdapter.Instance.DataCache.Rights.UserPrincipal, callback);
                                }
                            }
                        }
                        else
                        {
                            // Start copy to root node
                            DataAdapter.Instance.DataProvider.CopyToRootNode(pastecmisobjects, LastSelectedRootNode, DataAdapter.Instance.DataCache.Rights.UserPrincipal);
                        }
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        // TS 22.04.15 umbau auf mehrere
        //public bool CanPaste { get { return _CanPaste(DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_clipboard).Object_Selected); } }
        //private bool _CanPaste() { return _CanPaste(DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_clipboard).Object_Selected); }
        public bool CanPaste { get { return _CanPaste(DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_clipboard).Objects_Selected); } }

        private bool _CanPaste()
        {
            return _CanPaste(DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_clipboard).Objects_Selected);
        }

        // TS 22.04.15 umbau auf mehrere
        //private bool _CanPaste(IDocumentOrFolder sourceobj)
        private bool _CanPaste(List<IDocumentOrFolder> sourceobjectlist)
        {
            // TS 01.04.15 wenn paste im bereich clipboard aufgerufen wird dann weiterleiten
            // TS 20.04.16 bugfix: wenn bereits im workspace clipboard dann nicht weiterleiten sonst stackoverflow
            // if (Workspace == CSEnumProfileWorkspace.workspace_clipboard)
            if (Workspace == CSEnumProfileWorkspace.workspace_clipboard && !LastSelectedWorkspace.ToString().Equals(CSEnumProfileWorkspace.workspace_clipboard.ToString()))
                return DataAdapter.Instance.Processing(LastSelectedWorkspace).CanPaste;

            bool canpaste = DataAdapter.Instance.DataCache.ApplicationFullyInit;
            bool hasFile = false;
            bool hasFolder = false;
            IDocumentOrFolder parent = CreateObjectGetParent(null);
            canpaste = canpaste && !parent.isNotCreated;

            // TS 22.04.15 umbau auf mehrere
            foreach (IDocumentOrFolder sourceobj in sourceobjectlist)
            {                
                canpaste = canpaste && sourceobj.isClipboardObject;

                // Only allow file-operations on pure files (not mixed with folders)
                hasFolder = sourceobj.isFolder ? true : hasFolder;
                hasFile = sourceobj.isDocument ? true : hasFile;
                if(canpaste)
                {
                    if(hasFile)
                    {
                        canpaste = !hasFolder;
                    }
                }

                // nur die root-objekte im clipboard sind zugelassen für paste
                canpaste = canpaste && DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_clipboard).ObjectList.Contains(sourceobj);

                // Special behaviour for aktenplan-nodes
                if (LastSelectedWorkspace != CSEnumProfileWorkspace.workspace_aktenplan)
                {
                    // kann unterhalb von target erstellt werden ?
                    if (sourceobj.isFolder)
                    {
                        // TS 30.01.15
                        canpaste = canpaste && CanCreateFolder(sourceobj.structLevel);
                    }
                    else
                    {
                        // TS 18.12.13 ist für echtes paste nicht ok, da bei CanCreateDocuments auch erlaubt ist automatisch ein doklog anzulegen
                        //canpaste = canpaste && CanCreateDocuments;
                        // TS 06.02.14 jetzt doch, da für mfp import genau das gebraucht wird
                        //canpaste = canpaste && DataAdapter.Instance.DataCache.Objects(Workspace).Folder_Selected.canCreateDocument;
                        canpaste = canpaste && CanCreateDocuments;
                    }
                }else
                {
                    canpaste = canpaste && sourceobj.structLevel <= Statics.Constants.STRUCTLEVEL_09_DOKLOG;
                }
            }
            // TS 12.05.15
            //return canpaste;
            return canpaste && (sourceobjectlist != null && sourceobjectlist.Count > 0);
        }

        #endregion Paste (cmd_Paste)

        // ==================================================================
        #region PasteAsOriginal (cmd_PasteAsOriginal)

        public DelegateCommand cmd_PasteAsOriginal { get { return new DelegateCommand(PasteAsOriginal, _CanPasteAsOriginal); } }

        public void PasteAsOriginal_Hotkey() { PasteAsOriginal(); }
        private void PasteAsOriginal() { PasteAsOriginal(DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_clipboard).Objects_Selected, null); }

        public void PasteAsOriginal(List<IDocumentOrFolder> objectlist, IDocumentOrFolder target)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanPasteAsOriginal)
                {
                    if (Workspace == CSEnumProfileWorkspace.workspace_clipboard && !LastSelectedWorkspace.ToString().Equals(CSEnumProfileWorkspace.workspace_clipboard.ToString()))
                    {
                        DataAdapter.Instance.Processing(LastSelectedWorkspace).PasteAsOriginal(DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_clipboard).Objects_Selected, null);
                    }
                    else
                    {
                        List<cmisObjectType> pastecmisobjects = new List<cmisObjectType>();
                        foreach (IDocumentOrFolder singleobj in objectlist)
                        {
                            pastecmisobjects.Add(singleobj.CMISObject);
                        }

                        // Special behaviour for copy & PasteAsOriginal into rootnodes
                        if (LastSelectedWorkspace != CSEnumProfileWorkspace.workspace_aktenplan || LastSelectedRootNode == null)
                        {
                            bool autosave = DataAdapter.Instance.DataCache.Profile.Option_GetValueBoolean(CSEnumOptions.autosave_paste);
                            if (target == null)
                            {
                                target = DataAdapter.Instance.DataCache.Objects(Workspace).Folder_Selected;
                            }
                            if(target == null || target.objectId.Length == 0)
                            {
                                target = DataAdapter.Instance.DataCache.Objects(Workspace).Root;
                            }

                            cmisObjectType cmistarget = target.CMISObject;
                            IDocumentOrFolder obj = objectlist[objectlist.Count - 1];
                            CallbackAction callback = new CallbackAction(DataAdapter.Instance.Processing(target.Workspace).ObjectTransfer_Done, obj.isDocument);
                            if (!obj.isDocument)
                            {
                                DataAdapter.Instance.DataProvider.Paste(pastecmisobjects, cmistarget, autosave, true, DataAdapter.Instance.DataCache.Rights.UserPrincipal, callback);
                            }
                            else
                            {
                                if (target.structLevel == Constants.STRUCTLEVEL_09_DOKLOG)
                                {
                                    DataAdapter.Instance.DataProvider.Paste(pastecmisobjects, cmistarget, autosave, true, DataAdapter.Instance.DataCache.Rights.UserPrincipal, callback);
                                }
                                else
                                {
                                    cmisTypeContainer foldertypedefinition = DataAdapter.Instance.DataCache.Repository(Workspace).GetTypeContainerForStructLevel(Constants.STRUCTLEVEL_09_DOKLOG, target);
                                    DataAdapter.Instance.DataProvider.PasteDocumentsInFolders(foldertypedefinition, cmistarget, pastecmisobjects, autosave, DataAdapter.Instance.DataCache.Rights.UserPrincipal, callback);
                                }
                            }
                        }
                        else
                        {
                            // Start copy to root node
                            DataAdapter.Instance.DataProvider.CopyToRootNode(pastecmisobjects, LastSelectedRootNode, DataAdapter.Instance.DataCache.Rights.UserPrincipal);
                        }
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public bool CanPasteAsOriginal { get { return _CanPasteAsOriginal(DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_clipboard).Objects_Selected); } }

        private bool _CanPasteAsOriginal()
        {
            return _CanPasteAsOriginal(DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_clipboard).Objects_Selected);
        }

        private bool _CanPasteAsOriginal(List<IDocumentOrFolder> sourceobjectlist)
        {
            if (Workspace == CSEnumProfileWorkspace.workspace_clipboard && !LastSelectedWorkspace.ToString().Equals(CSEnumProfileWorkspace.workspace_clipboard.ToString()))
                return DataAdapter.Instance.Processing(LastSelectedWorkspace).CanPasteAsOriginal;

            bool canPasteAsOriginal = DataAdapter.Instance.DataCache.ApplicationFullyInit;
            IDocumentOrFolder parent = CreateObjectGetParent(null);
            canPasteAsOriginal = canPasteAsOriginal && !parent.isNotCreated;

            foreach (IDocumentOrFolder sourceobj in sourceobjectlist)
            {
                canPasteAsOriginal = canPasteAsOriginal && sourceobj.isClipboardObject;
                canPasteAsOriginal = canPasteAsOriginal && DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_clipboard).ObjectList.Contains(sourceobj);

                // Special behaviour for aktenplan-nodes
                if (LastSelectedWorkspace != CSEnumProfileWorkspace.workspace_aktenplan)
                {
                    if (sourceobj.isFolder)
                    {
                        canPasteAsOriginal = canPasteAsOriginal && CanCreateFolder(sourceobj.structLevel);
                    }
                    else
                    {
                        canPasteAsOriginal = canPasteAsOriginal && CanCreateDocuments;
                    }
                }
                else
                {
                    canPasteAsOriginal = canPasteAsOriginal && sourceobj.structLevel <= Statics.Constants.STRUCTLEVEL_09_DOKLOG;
                }
            }
            return canPasteAsOriginal && (sourceobjectlist != null && sourceobjectlist.Count > 0);
        }

        #endregion PasteAsOriginal (cmd_PasteAsOriginal)        

        // ==================================================================


        #region PasteAll (cmd_PasteAll)

        public DelegateCommand cmd_PasteAll { get { return new DelegateCommand(PasteAll, _CanPasteAll); } }

        private void PasteAll()
        {
            PasteAll(null);
        }

        public void PasteAll_Hotkey()
        {
            PasteAll(null);
        }

        /// <summary>
        /// CServer unterscheidet zwischen Copy oder Move und ruft diese
        /// Kein automatisches Speichern
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="target"></param>
        public void PasteAll(IDocumentOrFolder target)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanPasteAll)
                {
                    // TS 01.04.15 wenn paste im bereich clipboard aufgerufen wird dann weiterleiten
                    // TS 20.04.16 bugfix: wenn bereits im workspace clipboard dann nicht weiterleiten sonst stackoverflow
                    // if (Workspace == CSEnumProfileWorkspace.workspace_clipboard)
                    if (Workspace == CSEnumProfileWorkspace.workspace_clipboard && !LastSelectedWorkspace.ToString().Equals(CSEnumProfileWorkspace.workspace_clipboard.ToString()))
                        DataAdapter.Instance.Processing(LastSelectedWorkspace).PasteAll(null);
                    else
                    {
                        IDocumentOrFolder lastobj = null;
                        List<cmisObjectType> pastecmisobjects = new List<cmisObjectType>();
                        foreach (IDocumentOrFolder obj in DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_clipboard).ObjectList_TopLevel)
                        {
                            pastecmisobjects.Add(obj.CMISObject);
                            lastobj = obj;
                        }

                        // Special behaviour for copy & paste into rootnodes
                        if (LastSelectedWorkspace != CSEnumProfileWorkspace.workspace_aktenplan || LastSelectedRootNode != null)
                        {

                            bool autosave = DataAdapter.Instance.DataCache.Profile.Option_GetValueBoolean(CSEnumOptions.autosave_paste);
                            if (target == null)
                            {
                                target = DataAdapter.Instance.DataCache.Objects(Workspace).Folder_Selected;
                            }
                            if (target == null || target.objectId.Length == 0)
                            {
                                target = DataAdapter.Instance.DataCache.Objects(Workspace).Root;
                            }
                            cmisObjectType cmistarget = target.CMISObject;

                            // TS 08.02.16 zurückgebaut
                            CallbackAction callback = new CallbackAction(DataAdapter.Instance.Processing(target.Workspace).ObjectTransfer_Done, lastobj.isDocument);
                            if (!lastobj.isDocument)
                            {
                                DataAdapter.Instance.DataProvider.Paste(pastecmisobjects, cmistarget, autosave, DataAdapter.Instance.DataCache.Rights.UserPrincipal, callback);
                            }
                            else
                            {
                                if (target.structLevel == Constants.STRUCTLEVEL_09_DOKLOG)
                                {
                                    DataAdapter.Instance.DataProvider.Paste(pastecmisobjects, cmistarget, autosave, DataAdapter.Instance.DataCache.Rights.UserPrincipal, callback);
                                }
                                else
                                {
                                    cmisTypeContainer foldertypedefinition = DataAdapter.Instance.DataCache.Repository(Workspace).GetTypeContainerForStructLevel(Constants.STRUCTLEVEL_09_DOKLOG, target);
                                    DataAdapter.Instance.DataProvider.PasteDocumentsInFolders(foldertypedefinition, cmistarget, pastecmisobjects, autosave, DataAdapter.Instance.DataCache.Rights.UserPrincipal, callback);
                                }
                            }
                        }else
                        {
                            // Start copy to root node
                            DataAdapter.Instance.DataProvider.CopyToRootNode(pastecmisobjects, LastSelectedRootNode, DataAdapter.Instance.DataCache.Rights.UserPrincipal);
                        }
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        public bool CanPasteAll { get { return _CanPasteAll(); } }

        //private bool _CanPaste() { return _CanPaste(DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_clipboard).Object_Selected); }
        private bool _CanPasteAll()
        {
            // TS 01.04.15 wenn paste im bereich clipboard aufgerufen wird dann weiterleiten
            // TS 20.04.16 bugfix: wenn bereits im workspace clipboard dann nicht weiterleiten sonst stackoverflow
            // if (Workspace == CSEnumProfileWorkspace.workspace_clipboard)
            if (Workspace == CSEnumProfileWorkspace.workspace_clipboard && !LastSelectedWorkspace.ToString().Equals(CSEnumProfileWorkspace.workspace_clipboard.ToString()))
                return DataAdapter.Instance.Processing(LastSelectedWorkspace).CanPasteAll;

            bool canpaste = DataAdapter.Instance.DataCache.ApplicationFullyInit;

            // TS 18.03.16
            IDocumentOrFolder parent = CreateObjectGetParent(null);
            canpaste = canpaste && !parent.isNotCreated;

            canpaste = canpaste && DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_clipboard).ObjectList_TopLevel.Count > 0;

            // TS 22.06.15 nur wenn gleiches structlevel
            int structlevel = -1;

            foreach (IDocumentOrFolder sourceobj in DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_clipboard).ObjectList_TopLevel)
            {
                // TS 22.06.15
                if (structlevel == -1)
                    structlevel = sourceobj.structLevel;
                else
                    canpaste = canpaste && (structlevel == sourceobj.structLevel);

                canpaste = canpaste && sourceobj.isClipboardObject;
                // nur die root-objekte im clipboard sind zugelassen für paste
                canpaste = canpaste && DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_clipboard).ObjectList.Contains(sourceobj);

                // Special behaviour for aktenplan-nodes
                if (LastSelectedWorkspace != CSEnumProfileWorkspace.workspace_aktenplan)
                {

                    // kann unterhalb von target erstellt werden ?
                    if (sourceobj.isFolder)
                        canpaste = canpaste && CanCreateFolder(sourceobj.structLevel);
                    else
                    {
                        // TS 18.12.13 ist für echtes paste nicht ok, da bei CanCreateDocuments auch erlaubt ist automatisch ein doklog anzulegen
                        //canpaste = canpaste && CanCreateDocuments;
                        // TS 06.02.14 jetzt doch, da für mfp import genau das gebraucht wird
                        //canpaste = canpaste && DataAdapter.Instance.DataCache.Objects(Workspace).Folder_Selected.canCreateDocument;
                        canpaste = canpaste && CanCreateDocuments;
                    }
                }
                else
                {
                    canpaste = canpaste && sourceobj.structLevel <= Statics.Constants.STRUCTLEVEL_09_DOKLOG;
                }
            }
            return canpaste;
        }

        #endregion PasteAll (cmd_PasteAll)

        // ==================================================================

        #region Profile_Update (cmd_Profile_Update)

        public DelegateCommand cmd_Profile_Update { get { return new DelegateCommand(Profile_UpdateLayoutComponents, _CanProfile_UpdateLayoutComponents); } }

        public void Profile_UpdateLayoutComponents(string writelayoutid)
        {
            Profile_UpdateLayoutComponents(writelayoutid, null);
        }

        public void Profile_UpdateLayoutComponents(string writelayoutid, CallbackAction callback)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (_CanProfile_UpdateLayoutComponents(writelayoutid))
                {
                    string applicationid = DataAdapter.Instance.DataCache.Profile.ProfileApplication.id;
                    string currentlayoutid = DataAdapter.Instance.DataCache.Profile.ProfileLayout.id;
                    List<CServer.CSProfileComponent> profilecomponentinfo = DataAdapter.Instance.CollectProfileComponentInfo();

                    // wenn kein layoutname angegeben wird das aktuelle verwendet
                    // die prüfung ob dies editierbar ist findet bereits beim CanWriteProfile statt
                    if (writelayoutid == null) writelayoutid = "";
                    if (writelayoutid.Length == 0) writelayoutid = currentlayoutid;

                    // TS 04.09.12 nur wenn beschreibbare gefunden wurden
                    if (profilecomponentinfo.Count() > 0)
                        DataAdapter.Instance.DataProvider.Profile_UpdateLayoutComponents(DataAdapter.Instance.DataCache.Profile.UserProfile, applicationid, writelayoutid, profilecomponentinfo, DataAdapter.Instance.DataCache.Rights.UserPrincipal, callback);
                    // TS 30.04.15
                    else if (callback != null)
                        callback.Invoke();
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public void Profile_UpdateLayoutComponents()
        {
            Profile_UpdateLayoutComponents("");
        }

        // ------------------------------------------------------------------
        private bool _CanProfile_UpdateLayoutComponents(object parameter) { return _CanProfile_UpdateLayoutComponents((string)parameter); }

        private bool _CanProfile_UpdateLayoutComponents(string parameter)
        {
            if (parameter == null) parameter = "";
            return DataAdapter.Instance.DataCache.Profile.ProfileLayout != null
                && DataAdapter.Instance.DataCache.ApplicationFullyInit
                && (
                    (parameter.Length == 0 && DataAdapter.Instance.DataCache.Profile.ProfileLayout.editable) | (parameter.Length > 0)
                    );
        }

        #endregion Profile_Update (cmd_Profile_Update)

        // ==================================================================

        #region Profile_WriteProfile

        public void Profile_WriteProfile()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (_CanProfile_WriteProfile())
                {
                    DataAdapter.Instance.DataProvider.Profile_WriteUserProfile(DataAdapter.Instance.DataCache.Profile.UserProfile, DataAdapter.Instance.DataCache.Rights.UserPrincipal,new CallbackAction(Profile_WriteProfile_Done));
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        private void Profile_WriteProfile_Done()
        {
            if(DataAdapter.Instance.DataCache.ResponseStatus.success)
            {
                if (LocalStorage.IsLocalStorageInUse)
                {
                    SaveCache(true, false, false, false, false, false, null);
                }
            }
        }

        // ------------------------------------------------------------------
        // TS 10.05.12 immer auf true, da sonst beim start und z.b. auswahl und "setdefault" der anwendung nicht gespeichert werden kann
        //private bool _CanProfile_WriteProfile() { return DataAdapter.Instance.DataCache.ApplicationFullyInit; }
        // TS 15.10.14 wieder so gemacht wie es früher war, leider keine ahnung warum das geändert wurde
        //private bool _CanProfile_WriteProfile() { return DataAdapter.Instance.DataCache.ApplicationFullyInit; }
        private bool _CanProfile_WriteProfile() { return true; }

        #endregion Profile_WriteProfile

        // ==================================================================

        #region Query (cmd_Query)

        public DelegateCommand cmd_Query { get { return new DelegateCommand(_QueryCommand_Prerun, _CanQuery); } }

        // TS 17.07.17 nur test zunaechst, der ganze grosse mist rund um die adressen muss sowieso endlich raus hier, diese hunderten sonderlocken nerven so langsam
        public DelegateCommand cmd_QueryADR { get { return new DelegateCommand(_QueryCommand_Prerun_ADR, _CanQuery); } }

        /// <summary>
        /// Command-Prerunner for checking, which query shall be started
        /// </summary>
        private void _QueryCommand_Prerun()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            bool modeFound = false;
            try
            {
                // Prio 1 - Searchengine
                Dictionary<string, string> queryValues = DataAdapter.Instance.DataCache.Objects(Workspace).CollectQueryValues(QueryMode.QuerySearchEngine);
                modeFound = DefaultTextDispatcher.GetValuesWithoutDefaultText(queryValues).Count > 0;
                if (modeFound) { QuerySearchEngine(); }

                // Prio 2 - Fulltext standard
                if (!modeFound)
                {
                    queryValues = DataAdapter.Instance.DataCache.Objects(Workspace).CollectQueryValues(QueryMode.QueryFulltextAll);
                    modeFound = DefaultTextDispatcher.GetValuesWithoutDefaultText(queryValues).Count > 0;
                    if (modeFound) { QueryFulltextAll(); }
                }

                // Prio 3 - Meta
                if (!modeFound)
                { 
                    Query();
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // TS 17.07.117 den kram verdoppelt wegen der adressen
        private void _QueryCommand_Prerun_ADR()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            bool modeFound = false;
            try
            {
                // Prio 1 - Searchengine
                Dictionary<string, string> queryValues = DataAdapter.Instance.DataCache.Objects(Workspace).CollectQueryValues(QueryMode.QuerySearchEngine);
                modeFound = DefaultTextDispatcher.GetValuesWithoutDefaultText(queryValues).Count > 0;
                if (modeFound) { QuerySearchEngineADR(); }

                // Prio 2 - Fulltext standard
                if (!modeFound)
                {
                    queryValues = DataAdapter.Instance.DataCache.Objects(Workspace).CollectQueryValues(QueryMode.QueryFulltextAll);
                    modeFound = DefaultTextDispatcher.GetValuesWithoutDefaultText(queryValues).Count > 0;
                    if (modeFound) { QueryFulltextAll(); }
                }

                // Prio 3 - Meta
                if (!modeFound)
                {
                    _Query(QueryMode.QueryNew, CSEnumProfileWorkspace.workspace_adressen);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public void Query()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanQuery)
                {
                    _Query(QueryMode.QueryNew, Workspace);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // weiterer eingang (wird bisher nur verwendet aus EDesktop.QeryGesamt)
        public void Query(List<string> displayproperties, Dictionary<string, List<CSQueryToken>> repositoryquerytokens, Dictionary<string, List<CSOrderToken>> repositoryordertokens,
            bool autotruncate, bool getimagelist, bool getfulltextpreview, bool getparents, bool usesearchengine, CallbackAction callback)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanQuery)
                {
                    _QueryGiven(QueryMode.QueryNew, displayproperties, repositoryquerytokens, repositoryordertokens, autotruncate, getimagelist, getfulltextpreview, getparents, usesearchengine, Workspace, callback);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        //
        // ------------------------------------------------------------------
        public bool CanQuery { get { return _CanQuery(); } }

        public bool _CanQuery()
        {
            // TS 20.11.13
            if (this.Workspace.ToString().Equals(CSEnumProfileWorkspace.workspace_deleted.ToString()) || this.Workspace.ToString().Equals(CSEnumProfileWorkspace.workspace_lastedited.ToString()))
            {
                return DataAdapter.Instance.DataCache.ApplicationFullyInit;
            }
            else
            {
                // TS 10.12.13
                //DataAdapter.Instance.DataCache.Objects(Workspace).Folder_Selected.objectId == ""
                // && DataAdapter.Instance.DataCache.Objects(Workspace).ObjectList.Count > 1
                return (DataAdapter.Instance.DataCache.ApplicationFullyInit
                    && DataAdapter.Instance.DataCache.Objects(Workspace).ObjectList.Count > 1
                    && DataAdapter.Instance.DataCache.Objects(Workspace).ObjectList[1].isEmptyQueryObject);
            }
        }

        // ------------------------------------------------------------------

        #endregion Query (cmd_Query)

        // ==================================================================

        #region QueryFulltextAll (cmd_QueryFulltextAll)

        // TS 01.03.13: Volltextsuche hat keinen eigenen "echten" Command und muss dementsprechend im Code behandelt werden
        public DelegateCommand cmd_QueryFulltextAll { get { return new DelegateCommand(QueryFulltextAll, _CanQueryFulltextAll); } }

        public void QueryFulltextAll()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanQueryFulltextAll)
                {
                    // TS 22.01.14 wenn ws_statics gesetzt sind wird der jeweils aktuell gewählte ordner als parent für die suche (auch volltext) mitgegeben
                    bool found = false;
                    if (DataAdapter.Instance.DataCache.Objects(this.Workspace).StaticTypePresets.Count > 0)
                    {
                        IDocumentOrFolder selected = DataAdapter.Instance.DataCache.Objects(this.Workspace).Folder_Selected;
                        // TS 19.02.14 schleife nach oben
                        while (!found && !selected.objectId.Equals(DataAdapter.Instance.DataCache.Objects(this.Workspace).Root.objectId) && selected.objectId.Length > 0)
                        {
                            foreach (int level in DataAdapter.Instance.DataCache.Objects(this.Workspace).StaticTypePresets.Keys)
                            {
                                if (level == selected.structLevel)
                                {
                                    found = true;
                                    break;
                                }
                            }
                            // TS 19.02.14 schleife nach oben
                            if (!found)
                                selected = selected.Parent;
                        }
                        if (found)
                            _Query(QueryMode.QueryFulltextAll, selected.objectId, Workspace);
                    }
                    if (!found)
                        _Query(QueryMode.QueryFulltextAll, Workspace);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        public bool CanQueryFulltextAll { get { return _CanQueryFulltextAll(); } }

        private bool _CanQueryFulltextAll()
        {
            return (DataAdapter.Instance.DataCache.ApplicationFullyInit);
        }

        // ------------------------------------------------------------------

        #endregion QueryFulltextAll (cmd_QueryFulltextAll)

        // ==================================================================

        #region QuerySearchEngine (cmd_QuerySearchEngine)

        public DelegateCommand cmd_QuerySearchEngineADR { get { return new DelegateCommand(QuerySearchEngineADR, _CanQuerySearchEngine); } }

        public void QuerySearchEngineADR()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanQuerySearchEngine)
                {
                    // TS 11.06.15 trick: workspace_adressen wird abgefangen bei query und umgebogen auf ausschließlich default
                    _Query(QueryMode.QuerySearchEngine, CSEnumProfileWorkspace.workspace_adressen);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // TS 01.03.13: Volltextsuche hat keinen eigenen "echten" Command und muss dementsprechend im Code behandelt werden
        //public DelegateCommand cmd_QuerySearchEngineOverall { get { return new DelegateCommand(QuerySearchEngineOverall, _CanQuerySearchEngine); } }

        public DelegateCommand cmd_QuerySearchEngine { get { return new DelegateCommand(QuerySearchEngine, _CanQuerySearchEngine); } }

        //public void QuerySearchEngineOverall()
        //{
        //    if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
        //    try
        //    {
        //        if (CanQuerySearchEngine)
        //        {
        //            _Query(QueryMode.QuerySearchEngine, CSEnumProfileWorkspace.workspace_searchoverall);
        //            //_Query(QueryMode.QuerySearchEngine, CSEnumProfileWorkspace.workspace_default);
        //        }
        //    }
        //    catch (Exception e) { Log.Log.Error(e); }
        //    if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        //}

        public void QuerySearchEngine()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanQuerySearchEngine)
                {
                    if (DataAdapter.Instance.DataCache.Objects(Workspace).QuerySE_IsOverallQuery)
                    {
                        //overall = true;
                        // TS 14.11.16 die unterscheidung ist nicht mehr nötig
                        //_Query(QueryMode.QuerySearchEngine, CSEnumProfileWorkspace.workspace_searchoverall);
                        _Query(QueryMode.QuerySearchEngine, this.Workspace);
                    }
                    else
                    {
                        // TS 24.05.16
                        // TS 06.03.17 nicht mehr mit root vorbelegen da ansonsten die benutzerauswahl bei query_se_searchrepositories nicht berücksichtigt wird
                        // und stattdessen nur das repository des aktuellen workspaces verwendet wird
                        // string infolder = DataAdapter.Instance.DataCache.Objects(this.Workspace).Root.objectId;
                        string infolder = "";

                        // TS 22.01.14 wenn ws_statics gesetzt sind wird der jeweils aktuell gewählte ordner als parent für die suche (auch volltext) mitgegeben
                        bool found = false;
                        IDocumentOrFolder selected = DataAdapter.Instance.DataCache.Objects(this.Workspace).Folder_Selected;
                        if (DataAdapter.Instance.DataCache.Objects(this.Workspace).StaticTypePresets.Count > 0)
                        {
                            // TS 19.02.14 schleife nach oben
                            while (!found && !selected.objectId.Equals(DataAdapter.Instance.DataCache.Objects(this.Workspace).Root.objectId) && selected.objectId.Length > 0)
                            {
                                foreach (int level in DataAdapter.Instance.DataCache.Objects(this.Workspace).StaticTypePresets.Keys)
                                {
                                    if (level == selected.structLevel)
                                    {
                                        found = true;
                                        break;
                                    }
                                }
                                // TS 19.02.14 schleife nach oben
                                if (!found)
                                    selected = selected.Parent;
                            }
                            if (found)
                            {
                                infolder = selected.objectId;
                            }

                            // TS 04.05.17
                            // hier prüfen ob nicht: Statics.Constants.OBJECTID_FOLDER_DUMMY_MAIL_OTHERLOCATIONS weil in dem kann nicht gesucht werden
                            if (infolder.Equals(Statics.Constants.OBJECTID_FOLDER_DUMMY_MAIL_OTHERLOCATIONS))
                            {
                                string msg = LocalizationMapper.Instance["label_edp_mail_others_no_search"];
                                if (msg == null || msg.Length == 0) msg = "In diesem Ordner kann nicht gesucht werden";
                                DisplayWarnMessage(msg);
                                return;
                            }

                        }
                        // TS 24.05.16 auch ohne static presets den gewählten folder mitgeben
                        else
                        {
                            if (selected.objectId.Length > 0)
                                infolder = selected.objectId;
                        }

                        // Always fill the searchoverall-list on default-searches
                        // TS 14.11.16 die unterscheidung ist nicht mehr nötig
                        //if (Workspace == CSEnumProfileWorkspace.workspace_default)
                        //{
                        //    _Query(QueryMode.QuerySearchEngine, infolder, CSEnumProfileWorkspace.workspace_searchoverall);
                        //}
                        //else
                        //{
                        //    _Query(QueryMode.QuerySearchEngine, infolder, Workspace);
                        //}
                        _Query(QueryMode.QuerySearchEngine, infolder, Workspace);
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        public bool CanQuerySearchEngine { get { return _CanQuerySearchEngine(); } }

        private bool _CanQuerySearchEngine()
        {
            // TS 23.05.16
            //return (DataAdapter.Instance.DataCache.ApplicationFullyInit);
            bool ret = DataAdapter.Instance.DataCache.ApplicationFullyInit;
            if (ret) ret = DataAdapter.Instance.DataCache.Repository(this.Workspace).GetTypeContainerForStructLevel(Constants.STRUCTLEVEL_11_FULLTEXT) != null;
            //if (ret)
            //{
            //    ret = false;
            //    foreach (cmisTypeContainer container in DataAdapter.Instance.DataCache.Repository(this.Workspace).TypeDescendants)
            //    {
            //        IDocumentOrFolder obj = new DocumentOrFolder(container.type, this.Workspace);
            //        if (obj.structLevel == Constants.STRUCTLEVEL_11_FULLTEXT) ret = true;
            //    }
            //}
            return ret;
        }

        // ------------------------------------------------------------------

        #endregion QuerySearchEngine (cmd_QuerySearchEngine)

        // ==================================================================

        #region QueryMore

        //public DelegateCommand cmd_QueryNext { get { return new DelegateCommand(QueryNext, _CanQueryNext); } }
        public void QueryMore(CallbackAction callback, CallbackAction notDoneCallback)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanQueryMore)
                {
                    _Query(QueryMode.QueryMore, this.Workspace, callback, notDoneCallback);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        public bool CanQueryMore { get { return _CanQueryMore(); } }

        private bool _CanQueryMore()
        {
            return (DataAdapter.Instance.DataCache.ApplicationFullyInit
                     && DataAdapter.Instance.DataCache.Objects(Workspace).QueryResultCount > 0 && DataAdapter.Instance.DataCache.Objects(Workspace).QueryHasMoreResults);
        }

        // ------------------------------------------------------------------

        #endregion QueryMore

        // ==================================================================

        #region QueryNext (cmd_QueryNext)

        public DelegateCommand cmd_QueryNext { get { return new DelegateCommand(QueryNext, _CanQueryNext); } }

        public void QueryNext()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanQueryNext)
                {
                    _Query(QueryMode.QueryNext, this.Workspace);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        //private void _QueryNextP(object parameter) { QueryNext((CSEnumProfileWorkspace)parameter); }
        // ------------------------------------------------------------------
        public bool CanQueryNext { get { return _CanQueryNext(); } }

        private bool _CanQueryNext()
        {
            return (DataAdapter.Instance.DataCache.ApplicationFullyInit
                     && DataAdapter.Instance.DataCache.Objects(Workspace).QueryResultCount > 0 && DataAdapter.Instance.DataCache.Objects(Workspace).QueryHasMoreResults);
        }

        // ------------------------------------------------------------------

        #endregion QueryNext (cmd_QueryNext)

        // ==================================================================

        #region QueryPrev (cmd_QueryPrev)

        public DelegateCommand cmd_QueryPrev { get { return new DelegateCommand(QueryPrev, _CanQueryPrev); } }

        public void QueryPrev()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanQueryPrev)
                {
                    _Query(QueryMode.QueryPrev, Workspace);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        //private void _QueryPrevP(object parameter) { QueryPrev((CSEnumProfileWorkspace)parameter); }
        // ------------------------------------------------------------------
        public bool CanQueryPrev { get { return _CanQueryPrev(); } }

        private bool _CanQueryPrev()
        {
            // TS 20.03.14
            //            return (DataAdapter.Instance.DataCache.ApplicationFullyInit
            //         && DataAdapter.Instance.DataCache.Objects(Workspace).QueryResultCount > 0 && DataAdapter.Instance.DataCache.Objects(Workspace).QuerySkip > 0);
            return (DataAdapter.Instance.DataCache.ApplicationFullyInit
                     && DataAdapter.Instance.DataCache.Objects(Workspace).QueryResultCount > 0 && DataAdapter.Instance.DataCache.Objects(Workspace).QuerySkipStringList.Count > 1);
        }

        // ------------------------------------------------------------------

        #endregion QueryPrev (cmd_QueryPrev)

        // ==================================================================

        #region QueryDepending

        public void QueryDepending(string folderid)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanQueryDepending)
                {
                    ClearCache();
                    // TS 27.06.13 abholen von orderbyy presets
                    //List<CSOrderToken> givenorderby = DataAdapter.Instance.DataCache.Objects(Workspace).QueryOrderByPresets;
                    List<CServer.CSOrderToken> givenorderby = new List<CSOrderToken>();
                    _Query_GetSortDataFilter(ref givenorderby, false);

                    _QueryDepending(folderid, givenorderby, Workspace);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        public bool CanQueryDepending { get { return _CanQueryDepending(); } }

        private bool _CanQueryDepending()
        {
            return (DataAdapter.Instance.DataCache.ApplicationFullyInit
                && DataAdapter.Instance.DataCache.Objects(Workspace).Folder_Selected.objectId.Length > 0
                && DataAdapter.Instance.DataCache.Objects(Workspace).Folder_Selected.structLevel > 0);
        }

        // ------------------------------------------------------------------

        #endregion QueryDepending

        // ==================================================================

        #region QuerySingleObjectById

        //public DelegateCommand cmd_QuerySingleObjectById { get { return new DelegateCommand(QuerySingleObjectById, _CanQuerySingleObjectById); } }
        // TS 11.02.14 callback dazu
        //public void QuerySingleObjectById(string repository, string objectid, bool isdocument, CSEnumProfileWorkspace workspace)
        // TS 03.07.15
        // public void QuerySingleObjectById(string repository, string objectid, bool isdocument, CSEnumProfileWorkspace workspace, CallbackAction callback)        
        private List<CSQueryProperties> QueryCreatePropertiesList (string repository, string objectid, bool isdocument)
        {
            // Basics
            DataAdapter.Instance.DataCache.Objects(Workspace).QueryCollectionURL = new Uri("", UriKind.RelativeOrAbsolute);
            DataAdapter.Instance.DataCache.Objects(Workspace).QueryCollectionXML = "";
            bool autotruncate = false;
            int querylistsize = int.Parse(DataAdapter.Instance.DataCache.Profile.Option_GetValue(CSEnumOptions.querylistsize));
            int queryskipcount = 0;

            // Prepare QueryTokens
            List<CServer.CSQueryToken> querytokens = new List<CSQueryToken>();
            CSQueryToken token = new CSQueryToken();
            token.propertyname = "cmis:objectId";
            token.propertyvalue = objectid;
            token.propertytype = enumPropertyType.id;

            if (isdocument)
                token.propertyreptypeid = "cmis:document";
            else
                token.propertyreptypeid = "cmis:folder";
            querytokens.Add(token);

            // Get Sort order
            List<CServer.CSOrderToken> orderby = new List<CSOrderToken>();
            CSOrderToken firstorderby = new CSOrderToken();
            firstorderby.propertyname = "cmis:objectId";
            firstorderby.orderby = CSEnumOrderBy.asc;
            orderby.Add(firstorderby);

            // Get Properties
            List<CSQueryProperties> querypropertieslist = new List<CSQueryProperties>();
            CSQueryProperties queryproperties = _QueryCreateQueryPropertiesFromValues(repository, new List<string>(), querytokens, orderby, autotruncate, queryskipcount, false);
            querypropertieslist.Add(queryproperties);

            return querypropertieslist;
        }

        public void QuerySingleObjectById(string repository, string objectid, bool isdocument, bool setselection, CSEnumProfileWorkspace workspace, CallbackAction callback)
        {
            QuerySingleObjectById(repository, objectid, isdocument, setselection, false, workspace, callback);
        }

        public void QuerySingleObjectById(string repository, string objectid, bool isdocument, bool setselection, bool getparents, CSEnumProfileWorkspace workspace, CallbackAction callback)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (_CanQuerySingleObjectById)
                {
                    bool getimagelist = false;
                    bool getfulltextpreview = true;
                    bool usesearchengine = false;
                    int querylistsize = int.Parse(DataAdapter.Instance.DataCache.Profile.Option_GetValue(CSEnumOptions.querylistsize));

                    // TS 03.07.15
                    CallbackAction cb = null;
                    if (setselection)
                    {
                        // TS 11.02.14
                        cb = new CallbackAction(_QuerySingleObjectById_Done, objectid, null);
                        if (callback != null)
                            cb = new CallbackAction(_QuerySingleObjectById_Done, objectid, callback);
                    }
                    else if (callback != null)
                        cb = callback;

                    // TS 18.03.14 genereller umbau
                    List<CSQueryProperties> querypropertieslist = QueryCreatePropertiesList(repository, objectid, isdocument);

                    // TS 18.03.14 genereller umbau
                    DataAdapter.Instance.DataProvider.Query(querypropertieslist, new List<CSQueryProperties>(), getparents, getfulltextpreview, getimagelist, usesearchengine, querylistsize, workspace, false,
                                                            DataAdapter.Instance.DataCache.Rights.UserPrincipal, cb);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // TS 06.03.14
        //private void _QuerySingleObjectById_Done(CallbackAction callback)
        private void _QuerySingleObjectById_Done(string objectid, CallbackAction callback)
        {         
            // TS 06.03.14
            if (objectid != null && objectid.Length > 0 && !DataAdapter.Instance.DataCache.Objects(this.Workspace).Object_Selected.objectId.Equals(objectid))
            {
                this.SetSelectedObject(objectid);
            }

            // ruft als callback den standard Query_Done()
            _Query_Done(callback);
        }

        private bool _CanQuerySingleObjectById { get { return DataAdapter.Instance.DataCache.ApplicationFullyInit; } }

        #endregion QuerySingleObjectById

        // ==================================================================

        #region QueryStatics

        public void QueryStatics()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                List<cmisTypeContainer> staticlevels = new List<cmisTypeContainer>();
                Dictionary<int, cmisTypeContainer> staticpresets = DataAdapter.Instance.DataCache.Objects(Workspace).StaticTypePresets;
                foreach (KeyValuePair<int, cmisTypeContainer> kv in staticpresets)
                    staticlevels.Add(kv.Value);

                // TS 27.06.13 die StaticOrderByPresets werden noch nicht versorgt => siehe DataCache_Objects!!
                List<CSOrderToken> staticorderby = DataAdapter.Instance.DataCache.Objects(Workspace).StaticOrderByPresets;

                // TS 30.11.16 wenn keine StaticOrderByPresets angegeben sind dann die Sortierung aus dem TreeView verwenden
                // TODO: falls irgendwann mal aus ListView Statics getriggert werden muß das hier entsprechend angepaßt bzw. erewitert werden
                if (staticorderby.Count == 0)
                {
                    staticorderby = this.TreeOrderTokens;
                }

                QueryStatics(staticlevels, DataAdapter.Instance.DataCache.Objects(Workspace).StaticQueries, staticorderby);
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        //public void QueryStatics(List<int> structlevel)
        //{
        //    if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
        //    try
        //    {
        //        List<cmisTypeContainer> staticlevels = new List<cmisTypeContainer>();
        //        foreach (int level in structlevel)
        //        {
        //            cmisTypeContainer container = DataAdapter.Instance.DataCache.Repository(Workspace).GetTypeContainerForStructLevel(level);
        //            staticlevels.Add(container);
        //        }
        //        QueryStatics(staticlevels, null);
        //    }
        //    catch (Exception e) { Log.Log.Error(e); }
        //    if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        //}
        public void QueryStatics(List<cmisTypeContainer> staticlevels, Dictionary<string, string> staticqueries, List<CSOrderToken> staticorderby)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanQueryStatics)
                {
                    _QueryStatics(staticlevels, staticqueries, staticorderby, Workspace);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        public bool CanQueryStatics { get { return _CanQueryStatics(); } }

        private bool _CanQueryStatics()
        {
            return (DataAdapter.Instance.DataCache.ApplicationFullyInit);
        }

        // ------------------------------------------------------------------

        #endregion QueryStatics

        // ==================================================================

        #region QuerySetLast (cmd_QuerySetLast)

        public DelegateCommand cmd_QuerySetLast { get { return new DelegateCommand(QuerySetLast, _CanQuerySetLast); } }

        public void QuerySetLast()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                CSEnumProfileWorkspace targetWS = Workspace == CSEnumProfileWorkspace.workspace_searchoverall ? CSEnumProfileWorkspace.workspace_default : Workspace;
                if (CanQuerySetLast)
                {
                    if (DataAdapter.Instance.DataCache.Objects(targetWS).ObjectList[1].isEmptyQueryObject)
                    {
                        // TS 14.11.16
                        //DataAdapter.Instance.DataCache.Objects(Workspace).RestoreQueryValuesSQL();
                        if (DataAdapter.Instance.DataCache.Objects(targetWS).LastStoredQueryValuesWasSE)
                            DataAdapter.Instance.DataCache.Objects(targetWS).RestoreQueryValuesSE();
                        else
                            DataAdapter.Instance.DataCache.Objects(targetWS).RestoreQueryValuesSQL();
                    }
                    else
                    {
                        if(Workspace == CSEnumProfileWorkspace.workspace_searchoverall)
                        {
                            DataAdapter.Instance.Processing(CSEnumProfileWorkspace.workspace_default).ClearCache(false, new CallbackAction(QuerySetLast_ClearDone));
                        }                        
                    }
                    this.ClearCache(false, new CallbackAction(QuerySetLast_ClearDone));
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        private void QuerySetLast_ClearDone()
        {
            if (DataAdapter.Instance.DataCache.ResponseStatus.success)
            {
                CSEnumProfileWorkspace targetWS = Workspace == CSEnumProfileWorkspace.workspace_searchoverall ? CSEnumProfileWorkspace.workspace_default : Workspace;
                if (DataAdapter.Instance.DataCache.Objects(targetWS).LastStoredQueryValuesWasSE)
                    DataAdapter.Instance.DataCache.Objects(targetWS).RestoreQueryValuesSE();
                else
                    DataAdapter.Instance.DataCache.Objects(targetWS).RestoreQueryValuesSQL();
            }
        }

        // ------------------------------------------------------------------
        public bool CanQuerySetLast { get { return _CanQuerySetLast(); } }

        private bool _CanQuerySetLast()
        {
            CSEnumProfileWorkspace targetWS = Workspace == CSEnumProfileWorkspace.workspace_searchoverall ? CSEnumProfileWorkspace.workspace_default : Workspace;

            return (DataAdapter.Instance.DataCache.ApplicationFullyInit
                    && DataAdapter.Instance.DataCache.Objects(targetWS).ObjectList.Count > 1
                    && 
                    (DataAdapter.Instance.DataCache.Objects(targetWS).StoredQueryValuesSQL.Count > 0 || DataAdapter.Instance.DataCache.Objects(targetWS).StoredQueryValuesSE.Count > 0)
                    );
        }

        // ------------------------------------------------------------------

        #endregion QuerySetLast (cmd_QuerySetLast)

        // ==================================================================

        #region RecoverDeleted (cmd_RecoverDeleted)

        public DelegateCommand cmd_RecoverDeleted { get { return new DelegateCommand(RecoverDeleted, _CanRecoverDeleted); } }

        public void RecoverDeleted()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanRecoverDeleted)
                {
                    List<cmisObjectType> recoverlist = new List<cmisObjectType>();
                    foreach (IDocumentOrFolder obj in DataAdapter.Instance.DataCache.Objects(this.Workspace).Objects_Selected)
                    {
                        recoverlist.Add(obj.CMISObject);
                    }
                    if (recoverlist.Count > 0)
                    {
                        CallbackAction callback = new CallbackAction(Recover_Deleted_Done);
                        DataAdapter.Instance.DataProvider.RecoverDeleted(recoverlist, DataAdapter.Instance.DataCache.Rights.UserPrincipal, callback);
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public void Recover_Deleted_Done()
        {
            Refresh();
            DataAdapter.Instance.InformObservers();
        }

        // ------------------------------------------------------------------
        public bool CanRecoverDeleted { get { return _CanRecoverDeleted(); } }

        private bool _CanRecoverDeleted()
        {
            return DataAdapter.Instance.DataCache.Objects(Workspace).Object_Selected.objectId.Length > 0
                    && DataAdapter.Instance.DataCache.Objects(Workspace).Object_Selected.isNotCreated == false;
        }

        #endregion RecoverDeleted (cmd_RecoverDeleted)

        // ==================================================================

        #region Refresh (cmd_Refresh)

        public DelegateCommand cmd_Refresh { get { return new DelegateCommand(Refresh, _CanRefresh); } }

        public void Refresh()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanRefresh)
                {
                    // TS 22.01.14
                    //_Query(QueryMode.QueryRefresh, this.Workspace);
                    // TS 27.01.14
                    //if (DataAdapter.Instance.DataCache.Objects(this.Workspace).StaticTypePresets.Count > 0)
                    //    ClearCache();
                    //else
                    //    _Query(QueryMode.QueryRefresh, this.Workspace);
                    List<CSInformation> infoasked = new List<CSInformation>();
                    CSInformation info = new CSInformation();
                    info.informationid = CSEnumInformationId.Undefined;
                    switch (this.Workspace)
                    {
                        case CSEnumProfileWorkspace.workspace_aufgabe:
                            info.informationid = CSEnumInformationId.Current_AufgabeCount;
                            info.informationidSpecified = true;
                            break;

                        case CSEnumProfileWorkspace.workspace_mail:
                            info.informationid = CSEnumInformationId.Unread_MailCount;
                            info.informationidSpecified = true;
                            break;

                        case CSEnumProfileWorkspace.workspace_post:
                            info.informationid = CSEnumInformationId.Unread_PostCount;
                            info.informationidSpecified = true;
                            break;

                        case CSEnumProfileWorkspace.workspace_stapel:
                            info.informationid = CSEnumInformationId.Current_StapelCount;
                            info.informationidSpecified = true;
                            break;

                        case CSEnumProfileWorkspace.workspace_termin:
                            info.informationid = CSEnumInformationId.Current_TerminCount;
                            info.informationidSpecified = true;
                            break;
                    }
                    if (!info.informationid.ToString().Equals(CSEnumInformationId.Undefined.ToString()))
                    {
                        infoasked.Add(info);
                        DataAdapter.Instance.DataProvider.GetInformationList(infoasked, DataAdapter.Instance.DataCache.Rights.UserPrincipal, new CallbackAction(Refresh2));
                    }
                    else
                        Refresh2();
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        private void Refresh2()
        {
            // TS 10.10.14 wenn auf statischem objekt gerufen, dann dessen kinder entfernen und dann neu laden
            IDocumentOrFolder current = DataAdapter.Instance.DataCache.Objects(this.Workspace).Object_Selected;
            if (current.isDocument || current.structLevel == Constants.STRUCTLEVEL_09_DOKLOG)
            {
                current = current.Parent;
                if (current.structLevel == Constants.STRUCTLEVEL_09_DOKLOG)
                {
                    current = current.Parent;
                }
            }

            // wenn object is static
            // TS 12.11.14 zusätzlich prüfen ob es ein internalroot ist, dann soll "normales" refresh gemacht werden und nicht getchildren
            // if (DataAdapter.Instance.DataCache.Objects(this.Workspace).ObjectIdList_Statics.Contains(current.objectId))
            if (DataAdapter.Instance.DataCache.Objects(this.Workspace).ObjectIdList_Statics.Contains(current.objectId) && !DataAdapter.Instance.DataCache.Profile.Profile_IsInternalRootId(current.objectId))
            {
                // TS 12.11.14 das kann wohl raus weil es eh nicht mehr abgefragt wird
                //bool hasstaticchildren = false;
                //// wenn kinder vorhanden dann prüfen ob diese auch static sind
                //if (current.hasChildObjects && current.ChildObjects.Count > 0)
                //{
                //    // wenn das erste kind auch static ist dann den aktuellen ordner aus dem cache werfen und neu einlesen
                //    if (DataAdapter.Instance.DataCache.Objects(this.Workspace).ObjectIdList_Statics.Contains(current.ChildObjects[0].objectId))
                //    {
                //        hasstaticchildren = true;
                //    }
                //}

                //if (hasstaticchildren)
                //{
                //    // statische kinder vorhanden dann den ordner selbst neu einlesen
                //    // alles rekursiv sammeln (die kinder und kindeskinder nach oben) und aus dem cache werfen
                //    List<IDocumentOrFolder> descendants = DataAdapter.Instance.DataCache.Objects(this.Workspace).GetDescendants(current);
                //    List<string> descendantids = new List<string>();
                //    // jetzt die liste umdrehen (kinder nach oben)
                //    foreach (IDocumentOrFolder child in descendants)
                //    {
                //        descendantids.Insert(0, child.objectId);
                //    }
                //    // nun aus dem cache werfen
                //    DataAdapter.Instance.DataCache.Objects(this.Workspace).RemoveObjects(descendantids);
                //    // nun den ordner neu einlesen
                //    this.QuerySingleObjectById(current.RepositoryId, current.objectId, false, this.Workspace, null);
                //}
                //else
                //{
                //    // keine statischen kinder vorhanden dann den skipcount zurücksetzen, die kinder löschen und neu einlesen
                //    List<IDocumentOrFolder> descendants = DataAdapter.Instance.DataCache.Objects(this.Workspace).GetDescendants(current);
                //    List<string> descendantids = new List<string>();
                //    // jetzt die liste umdrehen (kinder nach oben)
                //    foreach (IDocumentOrFolder child in descendants)
                //    {
                //        descendantids.Insert(0, child.objectId);
                //    }
                //    // nun aus dem cache werfen
                //    DataAdapter.Instance.DataCache.Objects(this.Workspace).RemoveObjects(descendantids);

                //    current.ChildrenSkip = 0;
                //    GetChildren(current.objectId, new CallbackAction(SetSelectedObject, current.objectId));
                //}

                // alles rekursiv sammeln (die kinder und kindeskinder nach oben) und aus dem cache werfen
                List<IDocumentOrFolder> descendants = DataAdapter.Instance.DataCache.Objects(this.Workspace).GetDescendants(current);
                List<string> descendantids = new List<string>();
                // jetzt die liste umdrehen (kinder nach oben)
                foreach (IDocumentOrFolder child in descendants)
                {
                    descendantids.Insert(0, child.objectId);
                }
                // nicht gut dadurch landet er unten im tree
                // descendantids.Add(current.objectId);

                // auf den parent setzen
                SetSelectedObject(current.parentId);

                // nun aus dem cache werfen
                DataAdapter.Instance.DataCache.Objects(this.Workspace).RemoveObjects(descendantids);
                // nun den ordner neu einlesen
                CallbackAction callbackinner = new CallbackAction(GetChildren, current.objectId);
                CallbackAction callbackouter = new CallbackAction(SetSelectedObject, current.objectId, 0, callbackinner);

                //this.QuerySingleObjectById(current.RepositoryId, current.objectId, false, this.Workspace, new CallbackAction(GetChildren, current.objectId));
                this.QuerySingleObjectById(current.RepositoryId, current.objectId, false, true, this.Workspace, callbackouter);
            }
            else
            {
                // TS 29.04.14
                bool hasqueryfilters = false;
                CSProfileWorkspace ws = DataAdapter.Instance.DataCache.Profile.Profile_GetProfileWorkspaceFromType(this.Workspace);
                if (ws != null && ws.datafilters != null)
                {
                    foreach (CSProfileDataFilter df in ws.datafilters)
                    {
                        if (df.type == CSEnumProfileDataFilterType.ws_query)
                        {
                            hasqueryfilters = true;
                            break;
                        }
                    }
                }

                // TS 29.04.14
                //if (DataAdapter.Instance.DataCache.Objects(this.Workspace).StaticTypePresets.Count > 0 || DataAdapter.Instance.DataCache.Objects(this.Workspace).StaticQueries.Count > 0)
                if (!hasqueryfilters && (DataAdapter.Instance.DataCache.Objects(this.Workspace).StaticTypePresets.Count > 0 || DataAdapter.Instance.DataCache.Objects(this.Workspace).StaticQueries.Count > 0))
                    ClearCache();
                else
                {
                    // TS 21.03.14 skip leeren
                    DataAdapter.Instance.DataCache.Objects(this.Workspace).QuerySkipStringList.Clear();
                    _Query(QueryMode.QueryRefresh, this.Workspace);
                }
            }
        }

        // ------------------------------------------------------------------
        public bool CanRefresh { get { return _CanRefresh(); } }

        private bool _CanRefresh()
        {
            return DataAdapter.Instance.DataCache.ApplicationFullyInit;
        }

        // ------------------------------------------------------------------

        #endregion Refresh (cmd_Refresh)

        //RootNodeMarked

        // ==================================================================

        #region Save (cmd_Save)

        public DelegateCommand cmd_Save { get { return new DelegateCommand(Save, _CanSave); } }

        public void Save_Hotkey() { Save(null, null); }
        public void Save() { Save(null, null); }
        public void Save(CallbackAction callback)
        {
            Save(null, callback);
        }

        // TS 10.03.14
        //public void Save(List<IDocumentOrFolder> objectstosave)
        public void Save(List<IDocumentOrFolder> objectstosave, CallbackAction finalcallback)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanSave || objectstosave != null)
                {
                    List<cmisObjectType> cmisobjects = new List<cmisObjectType>();
                    List<cmisObjectType> newparents = new List<cmisObjectType>();
                    // TS 17.12.13
                    List<IDocumentOrFolder> savedobjects = new List<IDocumentOrFolder>();

                    // TS 01.07.15
                    string folderselectedid = DataAdapter.Instance.DataCache.Objects(Workspace).Folder_Selected.objectId;
                    // TS 25.11.16
                    string documentselectedid = DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected.objectId;
                    string setselectedid = "";

                    // TS 20.03.13
                    if (objectstosave == null)
                    {
                        foreach (IDocumentOrFolder obj in DataAdapter.Instance.DataCache.Objects(Workspace).ObjectList_InWork.View)
                        {
                            if (obj.isNotCreated || obj.isEdited || obj.canCheckIn || obj.isContentUpdated || obj.isCut || obj.MoveParentId.Length > 0)
                            {
                                cmisobjects.Add(obj.CMISObject);
                                IDocumentOrFolder newparent = DataAdapter.Instance.DataCache.Objects(Workspace).GetObjectById(obj.TempOrRealParentId);
                                newparents.Add(newparent.CMISObject);
                                savedobjects.Add(obj);
                                if (obj.isFolder && obj.objectId.Equals(folderselectedid))
                                    setselectedid = folderselectedid;
                                // TS 25.11.16
                                else if (obj.isDocument && obj.objectId.Equals(documentselectedid))
                                    setselectedid = documentselectedid;
                            }
                        }
                    }
                    // TS 20.03.13
                    else
                    {
                        foreach (IDocumentOrFolder obj in objectstosave)
                        {
                            if (obj.isNotCreated || obj.isEdited || obj.canCheckIn || obj.isContentUpdated || obj.isCut || obj.MoveParentId.Length > 0)
                            {
                                cmisobjects.Add(obj.CMISObject);
                                IDocumentOrFolder newparent = DataAdapter.Instance.DataCache.Objects(Workspace).GetObjectById(obj.TempOrRealParentId);
                                newparents.Add(newparent.CMISObject);
                                savedobjects.Add(obj);
                                if (obj.isFolder && obj.objectId.Equals(folderselectedid))
                                    setselectedid = folderselectedid;
                                // TS 25.11.16
                                else if (obj.isDocument && obj.objectId.Equals(documentselectedid))
                                    setselectedid = documentselectedid;
                            }
                        }
                    }

                    bool copyRelValidated = CheckSaveDataForCopyRelations(savedobjects);

                    if (copyRelValidated)
                    {
                        // TS 10.03.14 validate
                        bool ret = true;
                        string validatefailid = "";
                        try
                        {
                            foreach (IDocumentOrFolder obj in savedobjects)
                            {
                                validatefailid = obj.objectId;
                                if (ret) ret = obj.ValidateData();
                                if (!ret)
                                    break;
                            }
                        }
                        catch (Exception e) { Log.Log.Error(e); }
                        if (ret)
                        {
                            if (cmisobjects.Count() > 0)
                            {
                                // TS 10.03.14
                                //DataAdapter.Instance.DataProvider.Save(cmisobjects, newparents, DataAdapter.Instance.DataCache.Rights.UserPrincipal, new CallbackAction(_Save_Done, savedobjects));
                                // TS 01.07.15
                                //CallbackAction callback = new CallbackAction(_Save_Done, savedobjects, finalcallback);
                                // TS 02.07.15 wenn automatisch gespeichert wird z.b. durch wechsel auf anderes objekt
                                // und dann als callback bereits ein SetSelectedObject mitkommt dann kann das SetSelection im _Save_Done weggelassen werden
                                if (finalcallback != null && finalcallback.CallbackMethodName.StartsWith("SetSelectedObject"))
                                {
                                    setselectedid = null;
                                }

                                CallbackAction callback = new CallbackAction(_Save_Done, savedobjects, setselectedid, finalcallback);
                                DataAdapter.Instance.DataProvider.Save(cmisobjects, newparents, DataAdapter.Instance.DataCache.Rights.UserPrincipal, callback);
                            }
                            // TS 10.03.14
                            else
                            {
                                if (finalcallback != null)
                                    finalcallback.Invoke();
                            }
                        }
                        // wenn validierung fehlgeschlagen dann auf das fehlerhafte objekt positionieren
                        else
                        {
                            this.SetSelectedObject(validatefailid);
                        }
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        /// <summary>
        /// Checks the given objects whether at least one of it is a copy-relation AND is edited
        /// Or on Struct-Level 10 files, if the parent is a copy (whilst then it's content has been changed)
        /// </summary>
        /// <param name="checkObjects"></param>
        /// <returns></returns>
        private bool CheckSaveDataForCopyRelations(List<IDocumentOrFolder> checkObjects)
        {
            bool ret = false;
            bool relFound = false;
            List<string> copyObjIDs = new List<string>();
            List<IDocumentOrFolder> parentList = new List<IDocumentOrFolder>();

            foreach(IDocumentOrFolder obj in checkObjects)
            {
                // Check for the copy-object itself (all lvls)
                // TS 12.02.18
                // if((obj.isCopyRelation || obj[CSEnumCmisProperties.tmp_copyfrom_id] != null) && (obj.isEdited && !obj.isServerForcedSave))
                // TS 22.02.18 bei kopie per ablagemappe hat er bereits beim speichern geschimpft
                //if (
                //    (obj.isCopyRelation || obj[CSEnumCmisProperties.tmp_copyfrom_id] != null)
                //    && ((obj.isEdited && !obj.isServerForcedSave) || obj.isRealEdited || obj.hasChildrenInWork)
                //   )

                if (
                    (obj.isCopyRelation || obj[CSEnumCmisProperties.tmp_copyfrom_id] != null) 
                    && 
                    (
                    (obj.isEdited && !obj.isServerForcedSave) || obj.isRealEdited
                    )
                   )
                {
                    relFound = true;
                    copyObjIDs.Add(obj.objectId);
                }

                // Check for the objects-parent if the parent is a copy-object (only lvl 10)
                if (!relFound && obj.structLevel == Constants.STRUCTLEVEL_10_DOKUMENT)
                {
                    IDocumentOrFolder parentObject = DataAdapter.Instance.DataCache.Objects(obj.Workspace).GetObjectById(obj.parentId);
                    if(parentObject != null && parentObject.objectId.Length > 0 && parentObject.isCopyRelation && (obj.isEdited && !obj.isServerForcedSave))
                    {
                        relFound = true;
                        copyObjIDs.Add(parentObject.objectId);
                        parentList.Add(parentObject);
                    }

                }
            }

            // Show a validation-dialog, if necessary
            if(relFound == true)
            {
                if (showYesNoDialog(LocalizationMapper.Instance["msg_editcopy_validation"]))
                {
                    ret = true;
                    checkObjects.AddRange(parentList);

                    foreach (IDocumentOrFolder copyObj in checkObjects)
                    {
                        if (copyObjIDs.Contains(copyObj.objectId))
                        {
                            // Remove the copy-relations in sources
                            List<string> refreshObjects = copyObj.removeCopyRelations();

                            // Update the target-objects (if in cache)
                            foreach (string targetID in refreshObjects)
                            {
                                IDocumentOrFolder targetObject = DataAdapter.Instance.DataCache.Objects(this.Workspace).GetObjectById(targetID);
                                if (targetObject != null && targetObject.objectId.Length > 0)
                                {
                                    targetObject.removeRelationship(copyObj.objectId);
                                }
                            }

                            // Remove possible temp-values
                            if(copyObj[CSEnumCmisProperties.tmp_copyfrom_id] != null)
                            {
                                copyObj[CSEnumCmisProperties.tmp_copyfrom_id] = "";
                                copyObj[CSEnumCmisProperties.tmp_copyfrom_rep] = "";
                            }
                        }
                    }
                }
            }
            else
            {
                ret = true;
            }

            return ret;
        }

        // ------------------------------------------------------------------
        public bool CanSave { get { return _CanSave(); } }

        private bool _CanSave()
        {
            return (DataAdapter.Instance.DataCache.ApplicationFullyInit && DataAdapter.Instance.DataCache.Objects(Workspace).AnyInWork);
        }

        private void _Save_Done(object savedobjects, object selectedobjectid, CallbackAction finalcallback)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (DataAdapter.Instance.DataCache.ResponseStatus.success)
                {
                    // Trigger the viewer to lock the annotations
                    DataAdapter.Instance.DataCache.Info.Viewer_LockAnnotations = true;
                    DataAdapter.Instance.InformObservers(Workspace);

                    bool autosave_create = DataAdapter.Instance.DataCache.Profile.Option_GetValueBoolean(CSEnumOptions.autosave_create);
                    bool autosave_edit = DataAdapter.Instance.DataCache.Profile.Option_GetValueBoolean(CSEnumOptions.autosave_edit);

                    // ************************************
                    // TS 01.07.15 komplett anders versucht

                    //// TS 02.04.13 nicht einfach immer setzen sondern nur wenn KEIN objekt selektiert ist
                    ////this.SetSelectedObject(DataAdapter.Instance.DataCache.Objects(Workspace).ObjectList[DataAdapter.Instance.DataCache.Objects(Workspace).ObjectList.Count - 1].objectId);
                    //bool mustset = false;
                    //IDocumentOrFolder tmproot = DataAdapter.Instance.DataCache.Objects(Workspace).Root;
                    //IDocumentOrFolder tmplast = DataAdapter.Instance.DataCache.Objects(Workspace).ObjectList[DataAdapter.Instance.DataCache.Objects(Workspace).ObjectList.Count - 1];

                    //if (tmplast.isFolder && DataAdapter.Instance.DataCache.Objects(Workspace).Folder_Selected.objectId.Equals(tmproot.objectId))
                    //    mustset = true;
                    //// TS 17.04.13 so ist es falsch, da Document_Selected NIE auf Root sitzt (Folder schon) sondern auf einem leeren Dummy Objekt mit ID=""
                    ////else if (tmplast.isDocument && DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected.objectId.Equals(tmproot.objectId))
                    //else if (tmplast.isDocument && DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected.objectId.Length==0)
                    //    mustset = true;
                    //if (mustset)
                    //    this.SetSelectedObject(tmplast.objectId);
                    // ************************************

                    // TS 01.07.15 komplett anders versucht
                    // if (objectorlist != null && objectorlist.GetType().Name.StartsWith("List"))
                    string mustnotsetid = "";
                    // TS 13.11.15 was für ein bug !!
                    //if (selectedobjectid != null && selectedobjectid.GetType().Name.ToLower().Equals("String") && ((string)selectedobjectid).Length > 0)
                    if (selectedobjectid != null && selectedobjectid.GetType().Name.ToLower().Equals("string") && ((string)selectedobjectid).Length > 0)
                    {
                        // TS 25.11.16 setzen auch auf das letzte dokument: macht sinn bei neuen doks oder versionen
                        bool wasset = false;
                        string id = (string)selectedobjectid;
                        bool showfiles = true;

                        if (Workspace == CSEnumProfileWorkspace.workspace_default)
                        {
                            CSProfileComponent mainTree = DataAdapter.Instance.DataCache.Profile.Profile_GetComponentFromLayout(Constants.PROFILE_COMPONENTID_TREEVIEW);
                            if(mainTree != null)
                            {

                            }
                            showfiles = DataAdapter.Instance.DataCache.Profile.Option_GetValueBoolean(mainTree.options, CSEnumOptions.treeshowfiles);
                        }


                        // TS 25.11.16
                        IDocumentOrFolder docobj = DataAdapter.Instance.DataCache.Objects(Workspace).GetObjectById(DataAdapter.Instance.DataCache.Objects(Workspace).ObjectId_LastAddedDocument);
                        if ((id.Equals(docobj.objectId) || id.Equals(docobj[CSEnumCmisProperties.tmp_saveObjectIdUnsaved])) && showfiles)
                        {
                            if (!autosave_create && !autosave_edit)
                            {
                                this.SetSelectedObject(docobj.objectId);
                            }

                            // TS 23.06.17 auch das doklog neu laden wegen der collection
                            if (docobj.isDocument)
                            {
                                this.GetObjectsUpDown(docobj.ParentFolder, true, null);
                            }
                            mustnotsetid = docobj.objectId;
                            wasset = true;
                        }else
                        {
                            // Take care of subversioned documents and select the parentfolder afterwards
                            if(docobj.isDocument && docobj.ParentFolder != null && docobj.ParentFolder.objectId.Length > 0)
                            {
                                if (!autosave_create && !autosave_edit)
                                {
                                    this.SetSelectedObject(docobj.ParentFolder.objectId);
                                }

                                this.GetObjectsUpDown(docobj.ParentFolder, true, null);
                                mustnotsetid = docobj.objectId;
                                wasset = true;
                            }
                        }

                        // TS 25.11.16
                        if (!wasset)
                        {
                            foreach (string objid in DataAdapter.Instance.DataCache.Objects(Workspace).ObjectIdList_LastAddedFolders)
                            {
                                IDocumentOrFolder obj = DataAdapter.Instance.DataCache.Objects(Workspace).GetObjectById(objid);
                                if (id.Equals(objid) || id.Equals(obj[CSEnumCmisProperties.tmp_saveObjectIdUnsaved]))
                                {
                                    if (!autosave_create && !autosave_edit)
                                    {
                                        this.SetSelectedObject(objid);
                                    }
                                    mustnotsetid = objid;
                                    break;
                                }
                            }
                        }
                    }
                    if (((List<IDocumentOrFolder>)savedobjects).Count > 0)
                    {
                        // Get the saved objects (in the right cache) and refresh their DZCollection
                        List<IDocumentOrFolder> tmp_savedList = new List<IDocumentOrFolder>();
                        List<String> post2ecmObjects = new List<string>();
                        foreach(IDocumentOrFolder savedObj in (List<IDocumentOrFolder>)savedobjects)
                        {
                            tmp_savedList.Add(DataAdapter.Instance.DataCache.Objects(Workspace).GetObjectById(savedObj.objectId));

                            // Check for ECMDesktop-Relations and send back the acknowledgement
                            if (DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_default).ObjectIds_PreparedPostToECM.Count > 1 
                                && DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_default).ObjectIds_PreparedPostToECM.ContainsKey(savedObj.objectId) 
                                && !savedObj.isCut)
                            {
                                post2ecmObjects.Add("2C_" + savedObj.RepositoryId + "_" + savedObj.objectId);                                
                            }
                        }
                        RefreshDZCollectionsAfterProcess(tmp_savedList, "");

                        // Trigger the acknowledgement for the listening html-client for postobjects
                        if(post2ecmObjects.Count > 0)
                        {
                            DataAdapter.Instance.Processing(CSEnumProfileWorkspace.workspace_default).AcknowledgePostObjectsToECM(post2ecmObjects);
                        }

                        // Reset Editmode
                        DataAdapter.Instance.DataCache.Objects(Workspace).ObjectId_EditMode = "";
                        if (finalcallback != null)
                            finalcallback.Invoke();
                    }
                }
                else
                {
                    // Save data, that has been saved yet
                    if (((List<IDocumentOrFolder>)savedobjects).Count > 0)
                    {
                        // Get the saved objects (in the right cache) and refresh their DZCollection
                        List<IDocumentOrFolder> tmp_savedList = new List<IDocumentOrFolder>();
                        foreach (IDocumentOrFolder savedObj in (List<IDocumentOrFolder>)savedobjects)
                        {
                            tmp_savedList.Add(DataAdapter.Instance.DataCache.Objects(Workspace).GetObjectById(savedObj.objectId));
                        }
                        RefreshDZCollectionsAfterProcess(tmp_savedList, "");
                        DataAdapter.Instance.DataCache.Objects(Workspace).ObjectId_EditMode = "";
                        DataAdapter.Instance.InformObservers(Workspace);
                        if (finalcallback != null)
                            finalcallback.Invoke();
                    }

                    // Inform the user and maybe let him revert is changes
                    DispatcherTimer timer = new DispatcherTimer();
                    timer.Interval = new TimeSpan(0, 0, 0, 0, 20); // 100 ticks delay for updating UI
                    timer.Tick += delegate (object s, EventArgs ea)
                    {
                        timer.Stop();
                        timer.Tick -= delegate (object s1, EventArgs ea1) { };
                        timer = null;
                        // LocalizationMapper.Instance["msg_RequestForCancelAfterSaveFail"]
                        // TS 18.01.16
                        // if (showYesNoDialog(Localization.localstring.msg_RequestForCancelAfterSaveFail))
                        if (showYesNoDialog(LocalizationMapper.Instance["msg_RequestForCancelAfterSaveFail"]))
                        {
                            CancelAll();
                        }
                    };
                    timer.Start();
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------

        #endregion Save (cmd_Save)

        // ==================================================================

        #region setZdA

        public DelegateCommand cmd_EditZDA { get { return new DelegateCommand(EditZDA, _CanEditZDA); } }

        public void EditZDA()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanEditZDA)
                {
                    IDocumentOrFolder current = DataAdapter.Instance.DataCache.Objects(this.Workspace).Folder_Selected;

                    // status prüfen usw...
                    // dialog rufen
                    // anderes layout (rotes schild weg)
                    // save bzw. setzda im AdminContext rufen
                    // keine daten im rootnode gefunden dann den dialog rufen
                    View.Dialogs.dlgZDASettings zdaDialog = new View.Dialogs.dlgZDASettings(current, true);
                    DialogHandler.Show_Dialog(zdaDialog);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public bool CanEditZDA { get { return _CanEditZDA(); } }

        private bool _CanEditZDA()
        {
            bool ret = false;
            if (DataAdapter.Instance.DataCache.ApplicationFullyInit)
            {
                IDocumentOrFolder current = DataAdapter.Instance.DataCache.Objects(this.Workspace).Folder_Selected;
                string zdavalue = "";
                if (current.structLevel == 7)
                {
                    // 1001
                    zdavalue = current[CSEnumCmisProperties.AKTESYS_06];
                }
                else if (current.structLevel == 8)
                {                    
                    // fuer vorgaenge nur erlauben wenn nicht dieakte zda gesetzt ist denn es soll immer der oberste knoten bearbeitet werden
                    IDocumentOrFolder parent = current.Parent;
                    string parentzdavalue = parent[CSEnumCmisProperties.AKTESYS_06];
                    if (parentzdavalue == null || parentzdavalue.Length == 0 || !parentzdavalue.Equals(Statics.Constants.SPEC_PROPERTY_STATUS_ZDA))
                    {
                        zdavalue = current[CSEnumCmisProperties.VORGANGSYS_06];
                    }
                }
                else if(current.structLevel == 9)
                {
                    // fuer Dokumente nur erlauben wenn nicht die akte zda gesetzt ist und der Vorgang nicht zda gesetzt ist 
                    IDocumentOrFolder akte_Parent = current.Parent.Parent;
                    IDocumentOrFolder Vorgang_Parent = current.Parent;
                    string aktezdavalue = akte_Parent[CSEnumCmisProperties.AKTESYS_06];
                    string vorgangzdavalue = Vorgang_Parent[CSEnumCmisProperties.VORGANGSYS_06];

                    if ((aktezdavalue == null || aktezdavalue.Length == 0 || !aktezdavalue.Equals(Statics.Constants.SPEC_PROPERTY_STATUS_ZDA)) &&
                        (vorgangzdavalue == null || vorgangzdavalue.Length == 0 || !vorgangzdavalue.Equals(Statics.Constants.SPEC_PROPERTY_STATUS_ZDA)))
                    {
                        zdavalue = current[CSEnumCmisProperties.DOKUMENTSYS_06];
                    }
                }
                if (zdavalue != null && zdavalue.Length > 0)
                {
                    ret = zdavalue.Equals(Statics.Constants.SPEC_PROPERTY_STATUS_ZDA);
                }
            }
            return ret;
        }

        public void SetZDA(IDocumentOrFolder objectToSetZdA, bool reeditobject)
        {
            //if (CanSetZdA)
            //{
            List<IDocumentOrFolder> zdaStack = new List<IDocumentOrFolder>();
            List<string> zdaRefreshObjectsStack = new List<string>();

            // Add the zdA-object
            zdaStack.Add(objectToSetZdA);

            // Add the ID's of objects, which has to be refreshed after the zdA-setting
            zdaRefreshObjectsStack.Add(objectToSetZdA.objectId);
            foreach (IDocumentOrFolder desc in DataAdapter.Instance.DataCache.Objects(objectToSetZdA.Workspace).GetDescendants(objectToSetZdA))
            {
                zdaRefreshObjectsStack.Add(desc.objectId);
            }

            // Trigger setZDA or proceed normally
            if (zdaStack.Count > 0)
            {
                // HIER RAUSGENOMMEN
                CallbackAction callback = new CallbackAction(SetZdA_Done, zdaStack);
                DataAdapter.Instance.DataProvider.SetZdA(zdaStack, zdaRefreshObjectsStack, DataAdapter.Instance.DataCache.Rights.UserPrincipal, callback);
                // UND DAS HIER REIN
                //CallbackAction callback = new CallbackAction(SetZdA_Done, zdaStack);
                //DataAdapter.Instance.DataProvider.SetZdA(zdaStack, zdaRefreshObjectsStack, DataAdapter.Instance.DataCache.Rights.UserPrincipal, null);

                // TS 05.06.17 RAUS
                //List<IDocumentOrFolder> refList = new List<IDocumentOrFolder>();
                //refList.Add(objectToSetZdA);
                //refList.AddRange(DataAdapter.Instance.DataCache.Objects(objectToSetZdA.Workspace).GetDescendants(objectToSetZdA));
                //Cancel(refList, null);
                //DataAdapter.Instance.DataCache.Objects(Workspace).ObjectId_EditMode = "";
                //DataAdapter.Instance.InformObservers(Workspace);
            }
        }

        public void ShowZdA_Data(IDocumentOrFolder editedObject)
        {
            string repid = editedObject.RepositoryId;
            CSRootNode rootnode = DataAdapter.Instance.DataCache.RootNodes.GetNodeById(repid);
            Dictionary<string, string> fristlist = DataAdapter.Instance.DataCache.Meta.GetValuesAndMappingsListFromId("LIST_FRIST");
            string fristListVal;
            string fristCheckVal;
            // **********************************************************************

            string aufbewahrungsfrist = rootnode.aufbewahrungsfrist;
            if (aufbewahrungsfrist != null && aufbewahrungsfrist.Length > 0)
            {
                try
                {
                    foreach (KeyValuePair<string, string> entry in fristlist)
                    {
                        fristListVal = entry.Value.Substring(0, 2);
                        fristCheckVal = aufbewahrungsfrist.Contains("Jahre") ? aufbewahrungsfrist.Substring(0, aufbewahrungsfrist.IndexOf("Jahre") - 1) : aufbewahrungsfrist;

                        if (entry.Value.Equals(aufbewahrungsfrist) || entry.Key.Equals(fristCheckVal) || int.Parse(fristListVal) == int.Parse(fristCheckVal))
                        {
                            aufbewahrungsfrist = entry.Key;
                            break;
                        }
                    }
                }
                catch (Exception)
                {
                    aufbewahrungsfrist = "";
                }
            }
            // **********************************************************************

            string transferfrist = rootnode.transferfrist;
            if (transferfrist != null && transferfrist.Length > 0)
            {
                try
                {
                    foreach (KeyValuePair<string, string> entry in fristlist)
                    {
                        fristListVal = entry.Value.Substring(0, 2);
                        fristCheckVal = transferfrist.Contains("Jahre") ? transferfrist.Substring(0, transferfrist.IndexOf("Jahre") - 1) : transferfrist;
                        if (entry.Value.Equals(transferfrist) || entry.Key.Equals(fristCheckVal) || int.Parse(fristListVal) == int.Parse(fristCheckVal))
                        {
                            transferfrist = entry.Key;
                            break;
                        }
                    }
                }
                catch (Exception)
                {
                    transferfrist = "";
                }
            }
            // **********************************************************************

            string bewertung = rootnode.bewertung;
            // **********************************************************************

            // Recover Mapped-Value and set 'aussonderung
            string aussonderungsArt = "0";
            if (bewertung != null && bewertung.Length > 0)
            {
                Dictionary<string, string> bewertungslist = DataAdapter.Instance.DataCache.Meta.GetValuesAndMappingsListFromId("LIST_AUSSONDERUNGSART");
                foreach (KeyValuePair<string, string> entry in bewertungslist)
                {
                    if (entry.Equals(bewertung) || entry.Value.EndsWith(bewertung))
                    {
                        aussonderungsArt = entry.Key;
                        break;
                    }
                }
            }
            // **********************************************************************


            string zdaDatum = DateFormatHelper.GetBasicDateTime(DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(), DateTime.Now.Day.ToString(), "0", "0", "0");

            if (editedObject.structLevel == Statics.Constants.STRUCTLEVEL_07_AKTE)
            {
                //editedObject[CSEnumCmisProperties.AKTESYS_06] = 1001;
                editedObject[CSEnumCmisProperties.AKTESYS_19] = zdaDatum;
                editedObject[CSEnumCmisProperties.AKTESYS_15] = aufbewahrungsfrist;                    
                editedObject[CSEnumCmisProperties.AKTESYS_17] = aussonderungsArt;
                editedObject[CSEnumCmisProperties.AKTESYS_16] = transferfrist;
                //editedObject[CSEnumCmisProperties.AKTESYS_18] = bemerkung;
            }
            else if (editedObject.structLevel == Statics.Constants.STRUCTLEVEL_08_VORGANG)
            {
                //editedObject[CSEnumCmisProperties.VORGANGSYS_06] = 1001;
                editedObject[CSEnumCmisProperties.VORGANGSYS_19] = zdaDatum;
                editedObject[CSEnumCmisProperties.VORGANGSYS_15] = aufbewahrungsfrist;
                editedObject[CSEnumCmisProperties.VORGANGSYS_17] = aussonderungsArt;
                editedObject[CSEnumCmisProperties.VORGANGSYS_16] = transferfrist;
                //editedObject[CSEnumCmisProperties.VORGANGSYS_18] = bemerkung;
            } else if(editedObject.structLevel == Statics.Constants.STRUCTLEVEL_09_DOKLOG)
            {
                //editedObject[CSEnumCmisProperties.DOKUMENTSYS_06] = 1001;
                editedObject[CSEnumCmisProperties.DOKUMENTSYS_19] = zdaDatum;
                editedObject[CSEnumCmisProperties.DOKUMENTSYS_15] = aufbewahrungsfrist;
                editedObject[CSEnumCmisProperties.DOKUMENTSYS_17] = aussonderungsArt;
                editedObject[CSEnumCmisProperties.DOKUMENTSYS_16] = transferfrist;
                //editedObject[CSEnumCmisProperties.VORGANGSYS_18] = bemerkung;
            }

            View.Dialogs.dlgZDASettings zdaDialog = new View.Dialogs.dlgZDASettings(editedObject, false);
            DialogHandler.Show_Dialog(zdaDialog);

            // direkt speichern !!
            // TS 16.10.17 NICHT MEHR DIREKT SPEICHERN
            //SetZDA(editedObject, false);
        //}
        //    else
        //    {
        //        // keine daten im rootnode gefunden dann den dialog rufen
        //        View.Dialogs.dlgZDASettings zdaDialog = new View.Dialogs.dlgZDASettings(editedObject, false);
        //        DialogHandler.Show_Dialog(zdaDialog);
        //    }
        }

        // TS 16.10.17 das gibts so nicht mehr
        //public bool IsAutoZDA
        //{
        //    get
        //    {
        //        bool ret = false;
        //        IDocumentOrFolder current = DataAdapter.Instance.DataCache.Objects(this.Workspace).Folder_Selected;
        //        string repid = current.RepositoryId;
        //        CSRootNode rootnode = DataAdapter.Instance.DataCache.RootNodes.GetNodeById(repid);
        //        string aufbewahrungsfrist = rootnode.aufbewahrungsfrist;
        //        string bewertung = rootnode.bewertung;
        //        if (aufbewahrungsfrist != null && aufbewahrungsfrist.Length > 0 && bewertung != null && bewertung.Length > 0)
        //        {
        //            ret = true;
        //        }
        //        return ret;
        //    }
        //}

        //public bool CanSetZdA { get { return _CanSetZdA(); } }

        //private bool _CanSetZdA()
        //{
        //    bool ret = DataAdapter.Instance.DataCache.ApplicationFullyInit;
        //    if (ret)
        //    {
        //        //// TODO Prüfe Rechte hier oder so
        //        ret = DataAdapter.Instance.DataCache.Repository(this.Workspace).HasPermission(DataCache_Repository.RepositoryPermissions.PERMISSION_ZDA_6);
        //    }
        //    return ret;
        //}

        private void SetZdA_Done(object savedobjects)
        {
            try
            {
                List<IDocumentOrFolder> refObjects = (List<IDocumentOrFolder>)savedobjects;
                if (DataAdapter.Instance.DataCache.ResponseStatus.success)
                {
                    RefreshDZCollectionsAfterProcess(refObjects, "");
                    DataAdapter.Instance.DataCache.Objects(Workspace).ObjectId_EditMode = "";
                    DataAdapter.Instance.InformObservers(Workspace);
                }
            }
            catch (Exception){}
        }
        // ------------------------------------------------------------------
        #endregion

        #region setZdA

        public DelegateCommand cmd_UnsetZDA { get { return new DelegateCommand(UnsetZDA, _CanUnsetZDA); } }

        public void UnsetZDA()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanUnsetZDA)
                {
                    IDocumentOrFolder current = DataAdapter.Instance.DataCache.Objects(this.Workspace).Folder_Selected;
                    View.Dialogs.dlgUnsetZDASettings zdaDialog = new View.Dialogs.dlgUnsetZDASettings(current, true);
                    DialogHandler.Show_Dialog(zdaDialog);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public bool CanUnsetZDA { get { return _CanUnsetZDA(); } }

        private bool _CanUnsetZDA()
        {
            bool ret = false;
            if (DataAdapter.Instance.DataCache.ApplicationFullyInit)
            {
                IDocumentOrFolder current = DataAdapter.Instance.DataCache.Objects(this.Workspace).Folder_Selected;
                string zdavalue = "";
                if (current.structLevel == 7)
                {
                    // 1001
                    zdavalue = current[CSEnumCmisProperties.AKTESYS_06];
                }
                else if (current.structLevel == 8)
                {
                    // fuer vorgaenge nur erlauben wenn nicht die akte zda gesetzt ist denn es soll immer der oberste knoten bearbeitet werden
                    IDocumentOrFolder parent = current.Parent;
                    string parentzdavalue = parent[CSEnumCmisProperties.AKTESYS_06];
                    if (parentzdavalue == null || parentzdavalue.Length == 0 || !parentzdavalue.Equals(Statics.Constants.SPEC_PROPERTY_STATUS_ZDA))
                    {
                        zdavalue = current[CSEnumCmisProperties.VORGANGSYS_06];
                    }
                }
                else if (current.structLevel == 9)
                {
                    // fuer Dokumente nur erlauben wenn nicht die akte zda gesetzt ist und der Vorgang nicht zda gesetzt ist 
                    IDocumentOrFolder akte_Parent = current.Parent.Parent;
                    IDocumentOrFolder Vorgang_Parent = current.Parent;
                    string aktezdavalue = akte_Parent[CSEnumCmisProperties.AKTESYS_06];
                    string vorgangzdavalue = Vorgang_Parent[CSEnumCmisProperties.VORGANGSYS_06];

                    if ((aktezdavalue == null || aktezdavalue.Length == 0 || !aktezdavalue.Equals(Statics.Constants.SPEC_PROPERTY_STATUS_ZDA)) &&
                        (vorgangzdavalue == null || vorgangzdavalue.Length == 0 || !vorgangzdavalue.Equals(Statics.Constants.SPEC_PROPERTY_STATUS_ZDA)))
                    {
                        zdavalue = current[CSEnumCmisProperties.DOKUMENTSYS_06];
                    }
                }
                if (zdavalue != null && zdavalue.Length > 0)
                {
                    ret = zdavalue.Equals(Statics.Constants.SPEC_PROPERTY_STATUS_ZDA);
                }
            }
            return ret;
        }

        public void UnsetZDA(IDocumentOrFolder objectToUnsetZdA)
        {
            List<IDocumentOrFolder> zdaStack = new List<IDocumentOrFolder>();
            List<string> zdaRefreshObjectsStack = new List<string>();

            // Trigger setZDA or proceed normally
            CallbackAction callback = new CallbackAction(SetZdA_Done, objectToUnsetZdA);
            DataAdapter.Instance.DataProvider.UnsetZdA(objectToUnsetZdA, DataAdapter.Instance.DataCache.Rights.UserPrincipal, callback);
        }

        private void UnsetZdA_Done(object savedobject)
        {
            try
            {
                IDocumentOrFolder refObject = (IDocumentOrFolder)savedobject;
                if (DataAdapter.Instance.DataCache.ResponseStatus.success)
                {
                    RefreshDZCollectionsAfterProcess(refObject, "");
                    DataAdapter.Instance.DataCache.Objects(Workspace).ObjectId_EditMode = "";
                    DataAdapter.Instance.InformObservers(Workspace);
                }
            }
            catch (Exception) { }
        }

        // ------------------------------------------------------------------
        #endregion

        // ==================================================================

        #region SaveCache
        private bool _isQuotaSizeDialogHandled = false;
        private bool _isQuotaSizeIncreaseAllowed = false;
        public bool isQuotaSizeDialogHandled { get { return _isQuotaSizeDialogHandled; } set { _isQuotaSizeDialogHandled = value; } }
        public bool isQuotaSizeIncreaseAllowed { get { return _isQuotaSizeIncreaseAllowed; } set { _isQuotaSizeIncreaseAllowed = value; } }
        /// <summary>
        /// Saves all necessary cache-information into the localstorage (if possible) for later boot-upspeeding
        /// Is necessary for subpages and a fast booting of the application
        /// NOTE: This action has to run on the mainthread and takes several seconds!
        /// </summary>
        public void SaveCache( bool save_profile, bool save_meta, bool save_theme, bool save_rights, bool save_repository, bool save_roots, CallbackAction callback)
        {

            try
            {
                _isQuotaSizeDialogHandled = false;
                _isQuotaSizeIncreaseAllowed = false;

                LocalStorage.IncreaseInitialQuotaSize(this);

                DispatcherTimer timer = new System.Windows.Threading.DispatcherTimer();
                timer.Interval = new TimeSpan(0, 0, 0, 0, 250);
                timer.Tick += delegate (object s, EventArgs ea)
                {
                    if (_isQuotaSizeDialogHandled)
                    {
                    // Stahp
                    _isQuotaSizeDialogHandled = false;
                        timer.Stop();
                        timer = null;

                        if (_isQuotaSizeIncreaseAllowed)
                        {
                            // Save Profile-Cache
                            if (save_profile)
                            {
                                Log.Log.WriteToBrowserConsole("LocalStorage: Save Profile invoked");
                                string cacheData = ObjectSerializing.SerializeObject(DataAdapter.Instance.DataCache.Profile, true);
                                LocalStorage.SaveStringToLocalStorage(EnumHelper.GetValueFromCSEnum(Constants.EnumDataCachesInLocalStorage.LST_CACHE_DATA_PROFILE), cacheData);
                                CacheUserProfileLastModDate(null);
                            }

                            // Save Meta
                            if (save_meta)
                            {
                                Log.Log.WriteToBrowserConsole("LocalStorage: Save Meta invoked");
                                string cerealData_meta = ObjectSerializing.SerializeObject(DataAdapter.Instance.DataCache.ClientCom.CachedMetaData, true);
                                LocalStorage.SaveStringToLocalStorage(EnumHelper.GetValueFromCSEnum(Constants.EnumDataCachesInLocalStorage.LST_CACHE_DATA_META), cerealData_meta);
                                CacheMetaLastModHash(null);
                            }

                            // Save Themes
                            if (save_theme)
                            {
                                Log.Log.WriteToBrowserConsole("LocalStorage: Save Themes invoked");
                                string cerealData_theme = ObjectSerializing.SerializeObject(DataAdapter.Instance.DataCache.ClientCom.CachedThemeData, true);
                                LocalStorage.SaveStringToLocalStorage(EnumHelper.GetValueFromCSEnum(Constants.EnumDataCachesInLocalStorage.LST_CACHE_DATA_THEME), cerealData_theme);
                                CacheThemeLastModDate(null);
                                CacheIconLastModDate(null);
                            }

                            // Save Rights
                            if (save_rights)
                            {
                                Log.Log.WriteToBrowserConsole("LocalStorage: Save Rights invoked");
                                string cerealData_rights = ObjectSerializing.SerializeObject(DataAdapter.Instance.DataCache.Rights, true);
                                LocalStorage.SaveStringToLocalStorage(EnumHelper.GetValueFromCSEnum(Constants.EnumDataCachesInLocalStorage.LST_CACHE_DATA_RIGHTS), cerealData_rights);
                            }

                            // Save Repository
                            if (save_repository)
                            {
                                Log.Log.WriteToBrowserConsole("LocalStorage: Save Repository invoked");
                                string cerealData_repository = ObjectSerializing.SerializeObject(DataAdapter.Instance.DataCache.ClientCom.CachedRepositoryData, true);
                                LocalStorage.SaveStringToLocalStorage(EnumHelper.GetValueFromCSEnum(Constants.EnumDataCachesInLocalStorage.LST_CACHE_DATA_REPOSITORY), cerealData_repository);
                                CacheRepositoryLastModHash(null);
                            }

                            // Save Roots
                            if (save_roots)
                            {
                                Log.Log.WriteToBrowserConsole("LocalStorage: Save Roots invoked");
                                string cerealData_roots = ObjectSerializing.SerializeObject(DataAdapter.Instance.DataCache.ClientCom.CachedRootnodeData, true);
                                LocalStorage.SaveStringToLocalStorage(EnumHelper.GetValueFromCSEnum(Constants.EnumDataCachesInLocalStorage.LST_CACHE_DATA_ROOTS), cerealData_roots);
                                CacheRootsLastModHash(null);
                            }

                            // Invoke Callback
                            if (callback != null)
                            {
                                callback.Invoke();
                            }

                            // Close pending Dialogs
                            Deployment.Current.Dispatcher.BeginInvoke((ThreadStart)(() =>
                            {
                                LocalStorage.CloseWaitOnlyDialog();
                            }));
                        }
                    }

                };
                timer.Start();
            }
            catch (System.Exception ex) { if (ex.Message.Length > 0) { Log.Log.Debug(ex.Message); } }
        }

        public void SaveCache_RightsOnly()
        {
            // Save Rights
            string cerealData_rights = ObjectSerializing.SerializeObject(DataAdapter.Instance.DataCache.Rights, true);
            LocalStorage.SaveStringToLocalStorage(EnumHelper.GetValueFromCSEnum(Constants.EnumDataCachesInLocalStorage.LST_CACHE_DATA_RIGHTS), cerealData_rights);
        }

        /// <summary>
        /// Requests the userprofile-modification-date
        /// </summary>
        public void CacheUserProfileLastModDate(CallbackAction alternativeCallback)
        {
            CallbackAction callback = alternativeCallback == null ? new CallbackAction(CacheUserProfileLastModDate_Done) : alternativeCallback;
            DataAdapter.Instance.DataProvider.Profile_GetLastModificationDate(DataAdapter.Instance.DataCache.Rights.UserPrincipal, callback);
        }

        /// <summary>
        /// Requests the theme-modification-date
        /// If no alternative callback is given (== null), the webservice-response will be saved to the local storage
        /// </summary>
        public void CacheThemeLastModDate(CallbackAction alternativeCallback)
        {
            CallbackAction callback = alternativeCallback == null ? new CallbackAction(CacheThemeLastModDate_Done) : alternativeCallback;
            DataAdapter.Instance.DataProvider.Theme_GetLastModificationDate(DataAdapter.Instance.DataCache.Rights.UserPrincipal, callback);
        }

        /// <summary>
        /// Requests the icon-modification-date
        /// If no alternative callback is given (== null), the webservice-response will be saved to the local storage
        /// </summary>
        public void CacheIconLastModDate(CallbackAction alternativeCallback)
        {
            CallbackAction callback = alternativeCallback == null ? new CallbackAction(CacheIconLastModDate_Done) : alternativeCallback;
            DataAdapter.Instance.DataProvider.Icon_GetLastModificationDate(DataAdapter.Instance.DataCache.Rights.UserPrincipal, callback);
        }

        /// <summary>
        /// Requests the rights-modification-hash
        /// If no alternative callback is given (== null), the webservice-response will be saved to the local storage
        /// </summary>
        public void CacheRightsLastModHash(CallbackAction alternativeCallback)
        {
            CallbackAction callback =  alternativeCallback == null ? new CallbackAction(CacheRightsLastModHash_Done) : alternativeCallback;
            DataAdapter.Instance.DataProvider.Rights_GetCurrentHash(DataAdapter.Instance.DataCache.Rights.UserPrincipal, callback);
        }

        /// <summary>
        /// Requests the meta-modification-hash
        /// If no alternative callback is given (== null), the webservice-response will be saved to the local storage
        /// </summary>
        public void CacheMetaLastModHash(CallbackAction alternativeCallback)
        {
            string appidcurrent = DataAdapter.Instance.DataCache.Rights.ChoosenApplication.id;
            string repidcurrent = DataAdapter.Instance.DataCache.Rights.ChoosenRepository.repositoryId;
            CallbackAction callback = alternativeCallback == null ? new CallbackAction(CacheMetaLastModHash_Done) : alternativeCallback;
            DataAdapter.Instance.DataProvider.Meta_GetCurrentHash(appidcurrent, repidcurrent, DataAdapter.Instance.DataCache.Rights.UserPrincipal, callback);
        }

        /// <summary>
        /// Requests the rootnode-modification-hash
        /// If no alternative callback is given (== null), the webservice-response will be saved to the local storage
        /// </summary>
        public void CacheRootsLastModHash(CallbackAction alternativeCallback)
        {
            // Get allowed repositories
            List<string> allowedrepositories = new List<string>();
            foreach (cmisRepositoryEntryType reptype in DataAdapter.Instance.DataCache.Rights.UserRights.repositories)
            {
                allowedrepositories.Add(reptype.repositoryId);
            }

            CallbackAction callback = alternativeCallback == null ? new CallbackAction(CacheRootsLastModHash_Done) : alternativeCallback;
            DataAdapter.Instance.DataProvider.Roots_GetCurrentHash(allowedrepositories, DataAdapter.Instance.DataCache.Rights.UserPrincipal, callback);
        }

        /// <summary>
        /// Requests the icon-modification-date
        /// If no alternative callback is given (== null), the webservice-response will be saved to the local storage
        /// </summary>
        public void CacheRepositoryLastModHash(CallbackAction alternativeCallback)
        {
            string appidcurrent = DataAdapter.Instance.DataCache.Rights.ChoosenApplication.id;
            string repidcurrent = DataAdapter.Instance.DataCache.Rights.ChoosenRepository.repositoryId;
            List<string> reptries = new List<string>();
            List<CSEnumProfileWorkspace> workspaces = new List<CSEnumProfileWorkspace>();

            // Collect repository-lists and workspaces
            // For purpose watch "ReadRepositoryInfos()" in AppManager.cs
            reptries.Add(repidcurrent);
            workspaces.Add(CSEnumProfileWorkspace.workspace_default);
            reptries.Add(repidcurrent);
            workspaces.Add(CSEnumProfileWorkspace.workspace_default_related);
            List<string> reptriesDataCache = new List<string>();
            foreach (cmisRepositoryEntryType r in DataAdapter.Instance.DataCache.Rights.UserRights.repositories)
            {
                reptriesDataCache.Add(r.repositoryId);
            }
            foreach (CSProfileWorkspace profilews in DataAdapter.Instance.DataCache.Profile.ProfileApplication.workspaces)
            {
                if (!workspaces.Contains(profilews.typeid))
                {
                    if (profilews.repositories != null)
                    {
                        foreach (CSProfileRepository profilerep in profilews.repositories)
                        {
                            if (reptriesDataCache.Contains(profilerep.id) || profilerep.id.StartsWith(Statics.Constants.REPOSITORY_WORKFLOW.ToString() + "_"))
                            {
                                reptries.Add(profilerep.id);
                                workspaces.Add(profilews.typeid);
                            }
                        }
                    }
                    else
                    {
                        reptries.Add("0");
                        workspaces.Add(profilews.typeid);
                    }
                }
            }
            if (!workspaces.Contains(CSEnumProfileWorkspace.workspace_clipboard))
            {
                reptries.Add("0");
                workspaces.Add(CSEnumProfileWorkspace.workspace_clipboard);
            }
            if (!workspaces.Contains(CSEnumProfileWorkspace.workspace_aktenplan))
            {
                reptries.Add("0");
                workspaces.Add(CSEnumProfileWorkspace.workspace_aktenplan);
            }
            if (!workspaces.Contains(CSEnumProfileWorkspace.workspace_searchoverall))
            {
                reptries.Add("0");
                workspaces.Add(CSEnumProfileWorkspace.workspace_searchoverall);
            }

            CallbackAction callback = alternativeCallback == null ? new CallbackAction(CacheRepositoryLastModHash_Done) : alternativeCallback;
            DataAdapter.Instance.DataProvider.Repository_GetCurrentHash(appidcurrent, repidcurrent, reptries, workspaces, DataAdapter.Instance.DataCache.Rights.UserPrincipal, callback);
        }

        /// <summary>
        /// Saves the userprofile-modification-date
        /// </summary>
        private void CacheUserProfileLastModDate_Done()
        {
            if (DataAdapter.Instance.DataCache.ResponseStatus.success)
            {
                // Get Server_Timestamp
                string modDate_Server = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.Profile_LastModificationDate];
                Log.Log.WriteToBrowserConsole("LocalStorage: Save LastMod Profile invoked");
                LocalStorage.SaveStringToLocalStorage(EnumHelper.GetValueFromCSEnum(Constants.EnumDataCachesInLocalStorage.LST_CACHE_DATA_PROFILE) + "_LASTMOD_", modDate_Server);
            }
        }

        /// <summary>
        /// Saves the theme-modification-date
        /// </summary>
        private void CacheThemeLastModDate_Done()
        {
            if (DataAdapter.Instance.DataCache.ResponseStatus.success)
            {
                // Get Server_Timestamp
                string modDate_Server = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.Theme_LastModificationDate];
                Log.Log.WriteToBrowserConsole("LocalStorage: Save LastMod Theme invoked");
                LocalStorage.SaveStringToLocalStorage(EnumHelper.GetValueFromCSEnum(Constants.EnumDataCachesInLocalStorage.LST_CACHE_DATA_THEME) + "_LASTMOD_", modDate_Server);
            }
        }

        /// <summary>
        /// Saves the icon-modification-date
        /// </summary>
        private void CacheIconLastModDate_Done()
        {
            if (DataAdapter.Instance.DataCache.ResponseStatus.success)
            {
                // Get Server_Timestamp
                string modDate_Server = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.Icon_LastModificationDate];
                Log.Log.WriteToBrowserConsole("LocalStorage: Save LastMod Icon invoked");
                LocalStorage.SaveStringToLocalStorage(EnumHelper.GetValueFromCSEnum(Constants.EnumDataCachesInLocalStorage.LST_CACHE_DATA_ICON) + "_LASTMOD_", modDate_Server);
            }
        }

        /// <summary>
        /// Saves the rights-modification-hash
        /// </summary>
        private void CacheRightsLastModHash_Done()
        {
            if (DataAdapter.Instance.DataCache.ResponseStatus.success)
            {
                // Get Server_Hash
                string modHash_Server = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.Rights_ModificationHash];
                Log.Log.WriteToBrowserConsole("LocalStorage: Save LastMod Rights invoked");
                LocalStorage.SaveStringToLocalStorage(EnumHelper.GetValueFromCSEnum(Constants.EnumDataCachesInLocalStorage.LST_CACHE_DATA_RIGHTS) + "_LASTMOD_", modHash_Server);
            }
        }

        /// <summary>
        /// Saves the meta-modification-hash
        /// </summary>
        private void CacheMetaLastModHash_Done()
        {
            if (DataAdapter.Instance.DataCache.ResponseStatus.success)
            {
                // Get Server_Hash
                string modHash_Server = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.Meta_ModificationHash];
                Log.Log.WriteToBrowserConsole("LocalStorage: Save LastMod Meta invoked");
                LocalStorage.SaveStringToLocalStorage(EnumHelper.GetValueFromCSEnum(Constants.EnumDataCachesInLocalStorage.LST_CACHE_DATA_META) + "_LASTMOD_", modHash_Server);
            }
        }

        /// <summary>
        /// Saves the rights-modification-hash
        /// </summary>
        private void CacheRootsLastModHash_Done()
        {
            if (DataAdapter.Instance.DataCache.ResponseStatus.success)
            {
                // Get Server_Hash
                string modHash_Server = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.Roots_ModificationHash];
                Log.Log.WriteToBrowserConsole("LocalStorage: Save LastMod Roots invoked");
                LocalStorage.SaveStringToLocalStorage(EnumHelper.GetValueFromCSEnum(Constants.EnumDataCachesInLocalStorage.LST_CACHE_DATA_ROOTS) + "_LASTMOD_", modHash_Server);
            }
        }

        /// <summary>
        /// Saves the repository-modification-hash
        /// </summary>
        private void CacheRepositoryLastModHash_Done()
        {
            if (DataAdapter.Instance.DataCache.ResponseStatus.success)
            {
                // Get Server_Hash
                string modHash_Server = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.Repository_ModificationHash];
                Log.Log.WriteToBrowserConsole("LocalStorage: Save LastMod Repository invoked");
                LocalStorage.SaveStringToLocalStorage(EnumHelper.GetValueFromCSEnum(Constants.EnumDataCachesInLocalStorage.LST_CACHE_DATA_REPOSITORY) + "_LASTMOD_", modHash_Server);
            }
        }
        #endregion

        // ==================================================================

        #region setCheckoutState

        // TS 26.06.17 raus
        //public void PrepareChangeCheckoutState(IDocumentOrFolder objForSetting, bool isCheckedOut)
        //{
        //    // Save the object and set checkout on the webserver
        //    List<IDocumentOrFolder> saveList = new List<IDocumentOrFolder>();
        //    saveList.Add(objForSetting);
        //    CallbackAction callback = new CallbackAction(SetSingleCheckoutState, objForSetting, isCheckedOut);
        //    Save(saveList, callback);
        //}

        public void SetSingleCheckoutState(IDocumentOrFolder objForSetting, object isCheckedOut, object status_id)
        {
            List<IDocumentOrFolder> objList = new List<IDocumentOrFolder>();
            //objList.Add((IDocumentOrFolder)objForSetting);
            objList.Add(objForSetting);
            SetCheckoutState(objList, isCheckedOut, status_id, null);
        }

        public void SetCheckoutState(List<IDocumentOrFolder> objListForSetting, object isCheckedOut, object status_id, CallbackAction finalcallback)
        {
            bool checkedOut = (bool)isCheckedOut;
            string statusID = (string)status_id;
            //List<IDocumentOrFolder> objList = (List<IDocumentOrFolder>)objForSetting;            
            List<string> checkoutRefreshObjectsStack = new List<string>();
            List<IDocumentOrFolder> updateList = new List<IDocumentOrFolder>();

            foreach(IDocumentOrFolder singleObj in objListForSetting)
            {
                // Add the ID's of objects, which has to be refreshed after the checkout-setting
                updateList.Add(singleObj);
                foreach (IDocumentOrFolder desc in DataAdapter.Instance.DataCache.Objects(singleObj.Workspace).GetDescendants(singleObj))
                {
                    checkoutRefreshObjectsStack.Add(desc.objectId);
                    updateList.Add(desc);
                }
            }

            CallbackAction callback = new CallbackAction(SetCheckoutState_Done, updateList, finalcallback);
            DataAdapter.Instance.DataProvider.SetCheckoutState(objListForSetting, checkoutRefreshObjectsStack, checkedOut, statusID, DataAdapter.Instance.DataCache.Rights.UserPrincipal, callback);
        }

        private void SetCheckoutState_Done(object objSet, CallbackAction finalcallback)
        {
            List<IDocumentOrFolder> refObjects = (List<IDocumentOrFolder>)objSet;
            if (DataAdapter.Instance.DataCache.ResponseStatus.success)
            {
                // TS 26.06.17 aus zda uebernommen
                RefreshDZCollectionsAfterProcess(refObjects, "");

                DataAdapter.Instance.DataCache.Objects(Workspace).ObjectId_EditMode = "";
                DataAdapter.Instance.InformObservers(Workspace);
                if (finalcallback != null)
                {
                    finalcallback.Invoke();
                }
            }
            else
            {
                DisplayMessage(LocalizationMapper.Instance["msg_insufficient_rights"]);
                Cancel(refObjects, null);
            }
        }

        // ------------------------------------------------------------------
        #endregion

        // ==================================================================

        #region SetSelectedObject (cmd_SetSelectedObject)

        public DelegateCommand cmd_SetSelectedObject { get { return new DelegateCommand(_SetSelectedObject, _CanSetSelectedObject); } }

        // TS 22.12.14
        public void SetSelectedObject(string objectid, int callbacklevel, CallbackAction finalcallback)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                // TS 06.10.15
                if (_CanSetSelectedObject())
                {
                    IDocumentOrFolder tmp = DataAdapter.Instance.DataCache.Objects(Workspace).GetObjectById(objectid);
                    // TS 22.06.15
                    if (tmp != null && tmp.objectId.Length > 0)
                    {

                        // TS 10.02.17 im workspace_searchoverall muss nicht getobjectsupdown und so zeugs gelesen werden !!
                        // TS 22.03.17 im workspace_adressenselektion auch nicht
                        // if (this.Workspace == CSEnumProfileWorkspace.workspace_searchoverall)
                        if (this.Workspace == CSEnumProfileWorkspace.workspace_searchoverall || this.Workspace == CSEnumProfileWorkspace.workspace_adressenselektion)
                            callbacklevel = 20;

                        if (callbacklevel == 0)
                        {
                            if (Log.Log.IsDebugEnabled) Log.Log.Debug("Processing: SetSelectedObject: " + objectid);

                            // Send the selection-data to the adress-tool
                            SendECMSelection(tmp);

                            // Mark the object as read if it is not fully read
                            if (!tmp.MarkedAsReadOnSelection && tmp.isFolder)
                            {
                                List<IDocumentOrFolder> objList = new List<IDocumentOrFolder>();
                                objList.Add(tmp);
                                MarkObjectsAsRead(objList);
                            }
                            // Proceed selecting
                            string objectidlast = DataAdapter.Instance.DataCache.Objects(Workspace).Object_Selected.objectId;
                            // TS 01.07.15
                            List<IDocumentOrFolder> objectstosaveinstruct = null;
                            if (!tmp.isNotCreated && !tmp.isEdited)
                                objectstosaveinstruct = DataAdapter.Instance.DataCache.Objects(this.Workspace).GetObjectsInWorkForCurrentStruct(objectidlast, objectid);

                            if (objectstosaveinstruct == null || objectstosaveinstruct.Count == 0)
                                callbacklevel = 20;
                            else
                            {
                                // TS 14.04.14 es werden alle ungspeicherten zurückgeliefert und hier dann entweder automatisch gespeichert (autosave) oder eine meldung ausgegeben
                                bool autosave_create = DataAdapter.Instance.DataCache.Profile.Option_GetValueBoolean(CSEnumOptions.autosave_create);
                                bool autosave_edit = DataAdapter.Instance.DataCache.Profile.Option_GetValueBoolean(CSEnumOptions.autosave_edit);
                                List<IDocumentOrFolder> autosaves = new List<IDocumentOrFolder>();
                                List<IDocumentOrFolder> notautosaves = new List<IDocumentOrFolder>();
                                foreach (IDocumentOrFolder obj in objectstosaveinstruct)
                                {
                                    if ((obj.isNotCreated && autosave_create) || (obj.isEdited && autosave_edit))
                                        autosaves.Add(obj);
                                    else
                                        notautosaves.Add(obj);
                                }
                                if (autosaves.Count > 0)
                                {
                                    // automatisch speichern
                                    bool ret = true;
                                    string validatefailid = "";
                                    try
                                    {
                                        foreach (IDocumentOrFolder obj in autosaves)
                                        {
                                            validatefailid = obj.objectId;
                                            if (ret) ret = obj.ValidateData();
                                            if (!ret)
                                                break;
                                        }
                                    }
                                    catch (Exception e) { Log.Log.Error(e); }
                                    if (ret)
                                    {
                                        this.Save(autosaves, new CallbackAction(SetSelectedObject, objectid, 10, finalcallback));
                                    }
                                    else
                                    {
                                        objectid = validatefailid;
                                        callbacklevel = 20;
                                    }
                                }
                                if (notautosaves.Count > 0)
                                {
                                    callbacklevel = 20;
                                }
                            }
                        }
                        if (callbacklevel == 10)
                        {
                            if (DataAdapter.Instance.DataCache.ResponseStatus.success)
                            {
                                callbacklevel = 20;
                            }
                        }

                        if (callbacklevel == 20)
                        {
                            // TS 11.09.12 gleich setzen damit nicht z.b. listview zwischendurch auf den alten hüpft (ausgelöst durch updateui z.b. von responsestatus)
                            DataAdapter.Instance.DataCache.Objects(Workspace).SetSelectedObject(objectid);

                            // TS 20.05.16 umgehängt auf Processing.SetSelectedObject
                            // Update Notification-State in InfoCache
                            if (tmp.structLevel < 9)
                            {
                                DataAdapter.Instance.DataCache.Info.UpdateAlertNotification(tmp.objectId, tmp.Workspace, false);
                            }

                            // TS 10.02.17 im searchoverall_workspace muss nicht getobjectsupdown und so zeugs gelesen werden !!                            
                            if (this.Workspace == CSEnumProfileWorkspace.workspace_searchoverall)
                            {
                                callbacklevel = 100;
                            }
                            // TS 22.03.17 in workspace_adressenselektion auch nicht alles lesen
                            else if (this.Workspace == CSEnumProfileWorkspace.workspace_adressenselektion)
                            {
                                callbacklevel = 97;
                            }
                            else
                            {
                                // TS 22.03.17 raus da nicht benötigt
                                //if (Workspace != CSEnumProfileWorkspace.workspace_clipboard && Workspace != CSEnumProfileWorkspace.workspace_adressen && !Workspace.ToString().EndsWith("_related"))
                                //{
                                //    if (DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_adressen).ExistsObject(objectid))
                                //    {
                                //        DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_adressen).SetSelectedObject(objectid);
                                //    }
                                //}

                                // TS 26.05.15 wenn adressen, dann auch die kinder nachlesen da nun in eigenem workspace und anderes verhalten
                                // erstmal wieder rausgenommen
                                bool forcereadchildren = false;
                                //if (Workspace == CSEnumProfileWorkspace.workspace_adressen)
                                //    forcereadchildren = true;
                                // TS 20.01.16
                                if (!GetObjectsUpDown(tmp, forcereadchildren, new CallbackAction(SetSelectedObject, objectid, 97, finalcallback)))
                                    callbacklevel = 97;
                            }
                        }
                        // TS 13.04.17 ohje das war komplett verrutscht !!!
                        if (callbacklevel == 97)
                        {
                            if (!tmp.isInitFully && !tmp.isDocument
                                && !tmp.Workspace.ToString().StartsWith(CSEnumProfileWorkspace.workspace_clipboard.ToString())
                                && !tmp.Workspace.ToString().StartsWith(CSEnumProfileWorkspace.workspace_deleted.ToString())
                                && !tmp.Workspace.ToString().StartsWith(CSEnumProfileWorkspace.workspace_workflow.ToString())
                                && !tmp.Workspace.ToString().StartsWith(CSEnumProfileWorkspace.workspace_lastedited.ToString()))
                            {
                                DataAdapter.Instance.DataProvider.GetObject(tmp.RepositoryId, tmp.objectId, tmp.Workspace, DataAdapter.Instance.DataCache.Rights.UserPrincipal,
                                    new CallbackAction(SetSelectedObject, objectid, 98, finalcallback));
                            }
                            else
                            {
                                callbacklevel = 98;
                            }
                        }

                        // TS 23.05.13 annotations nachladen
                        if (callbacklevel == 98)
                        {
                            bool processed = false;
                            IDocumentOrFolder tmpdoc = DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected;
                            foreach (IDocumentOrFolder child in tmpdoc.ChildObjects)
                            {
                                if (child.CopyType == CSEnumDocumentCopyTypes.COPYTYPE_OVERLAY)
                                {
                                    if (child.AnnotationFile == null)
                                    {
                                        DataAdapter.Instance.DataProvider.GetAnnotationFile(child.CMISObject, DataAdapter.Instance.DataCache.Rights.UserPrincipal,
                                            new CallbackAction(SetSelectedObject, objectid, 99, finalcallback));
                                        processed = true;
                                    }
                                    break;
                                }
                            }
                            if (!processed)
                                callbacklevel = 99;
                        }

                        // fortsetzung und ende ...
                        if (callbacklevel == 99)
                        {
                            // TS 01.04.15 für die ablagemappe
                            if (Workspace != CSEnumProfileWorkspace.workspace_clipboard && !Workspace.ToString().EndsWith("_related"))
                                LastSelectedWorkspace = Workspace;

                            // TS 02.02.17 wenn dokument geklickt und adressen oder desktop offen und docked sind, dann meldung
                            if (this.Workspace == CSEnumProfileWorkspace.workspace_default)
                            {
                                IDocumentOrFolder tmpdoc = DataAdapter.Instance.DataCache.Objects(Workspace).Document_Selected;
                                if (tmpdoc.isDocument && tmpdoc.DownloadName != null && tmpdoc.DownloadName.Length > 0 && !tmpdoc.Parent.isAdressenDokument)
                                {
                                    CSProfileComponent pc = DataAdapter.Instance.DataCache.Profile.Profile_GetComponentFromLayout(CSEnumProfileComponentType.ADRESSEN, Constants.PROFILE_COMPONENTID_ADRESSEN);
                                    if (pc != null && pc.docked && ((MainPage)App.Current.RootVisual).IsViewAdressen)
                                        DisplayWarnMessage(LocalizationMapper.Instance["msg_ShouldHideAdressen"]);
                                    pc = DataAdapter.Instance.DataCache.Profile.Profile_GetComponentFromLayout(CSEnumProfileComponentType.EDESKTOP, Constants.PROFILE_COMPONENTID_EDESKTOP);
                                    if (pc != null && pc.docked && ((MainPage)App.Current.RootVisual).IsViewDesktop)
                                        DisplayWarnMessage(LocalizationMapper.Instance["msg_ShouldHideDesktop"]);
                                }
                            }                            

                            DataAdapter.Instance.InformObservers();
                            if (finalcallback != null) finalcallback.Invoke();

                            // Prepare Selection of the first fulltext-match
                            if (tmp.FulltextCoordinates.Count > 0)
                            {
                                SelectFirstFulltextMatchPage(tmp);
                            }
                        }

                        // TS 10.02.17
                        if (callbacklevel == 100)
                        {
                            if (finalcallback != null) finalcallback.Invoke();                          
                        }

                        tmp = null;
                    }
                }
            }
            //catch (System.ComponentModel.DataAnnotations.ValidationException ve) { throw ve; }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public void SetSelectedObject(string objectid)
        {
            SetSelectedObject(objectid, 0, null);
        }

        private void SelectFirstFulltextMatchPage(IDocumentOrFolder doc)
        {
            // Set the first matching page
            foreach (KeyValuePair<string, PointCollection> fulltext_entry in doc.FulltextCoordinates)
            {
                int firstPos = fulltext_entry.Key.IndexOf("|~|");
                int lastPos = fulltext_entry.Key.LastIndexOf("|~|");
                int coordPage = int.Parse(fulltext_entry.Key.Substring(firstPos + 3, lastPos - firstPos - 3));
                string docID = fulltext_entry.Key.Substring(0, fulltext_entry.Key.IndexOf("|~|"));
                IDocumentOrFolder selDoc = DataAdapter.Instance.DataCache.Objects(Workspace).GetObjectById(docID);

                // TS 22.12.17 nur noch treffer auf originale und inhaltliche versionen zulassen bis das von der volltextsuche mal ordentlich kommt
                if (selDoc.isValidFulltextHighlightDocument)
                {
                    // TS 22.12.17 nur auf seite navigieren wenn die auch bereits konvertiert ist
                    if (coordPage >= selDoc.DZPageCount && coordPage >= selDoc.DZPageCountConverted)
                    {
                        selDoc.SelectedPage = coordPage - 1;
                    }
                    SetSelectedObject(docID);
                    // TS 22.12.17 den break aus der schleife nur machen wenn er ein sinnvolles dokument gefunden hat
                    // weil evtl. zunaechst von volltext mehrere treffer auf unterschiedliche versionen kommen
                    // und wir dann halt solange schleifen bis eine anzeigbare dabei ist
                    break;
                }
                // TS 22.12.17 den break aus der schleife nur machen wenn er ein sinnvolles dokument gefunden hat (s.o.)
                //break;
            }
        }

        /// <summary>
        /// Send RCPush-Selectio to the adress-tool
        /// </summary>
        /// <param name="selection"></param>
        private void SendECMSelection(IDocumentOrFolder selection)
        {            
            // Create the 2C-Display-ID
            string link = "2C_" + selection.RepositoryId + "_" + selection.objectId;
            List<string> dummyList = new List<string>();
            dummyList.Add(link);

            // Create the RoutingInfo
            CSRCPushPullUIRouting routinginfo = new CSRCPushPullUIRouting();
            CSEnumRCPushUIListeners?[] receiver = new CSEnumRCPushUIListeners?[2];
            receiver[0] = CSEnumRCPushUIListeners.HTML_CLIENT_ADRESS;
            receiver[1] = CSEnumRCPushUIListeners.HTML_CLIENT_DESKTOP;
            routinginfo.sendto_componentorvirtualids = receiver;

            // Send a PushUI
            DataAdapter.Instance.DataProvider.RCPush_UI(routinginfo, CSEnumRCPushCommands.ecmselection, dummyList, null, DataAdapter.Instance.DataCache.Rights.UserPrincipal);
        }

        // TS 23.12.14
        private void _SetSelectedObject(object objectid) { SetSelectedObject((string)objectid, 0, null); }

        // ------------------------------------------------------------------
        // TS 06.10.15 raus weil sonst doppelt (s.u.)
        //public bool CanSetSelectedObject { get { return _CanSetSelectedObject(); } }
        // TS 06.10.15 die komplette prüfung von oben hier rein damit tree und list das vor dem wechsel abfragen können
        public bool CanSetSelectedObject(string objectid, bool displaywarning)
        {
            bool ret = false;
            IDocumentOrFolder tmp = DataAdapter.Instance.DataCache.Objects(Workspace).GetObjectById(objectid);
            if (tmp != null && tmp.objectId.Length > 0)
            {
                string objectidlast = DataAdapter.Instance.DataCache.Objects(Workspace).Object_Selected.objectId;

                List<IDocumentOrFolder> objectstosaveinstruct = null;
                if (!tmp.isNotCreated && !tmp.isEdited)
                    objectstosaveinstruct = DataAdapter.Instance.DataCache.Objects(this.Workspace).GetObjectsInWorkForCurrentStruct(objectidlast, objectid);

                if (objectstosaveinstruct == null || objectstosaveinstruct.Count == 0)
                    ret = true;
                else
                {
                    // TS 14.04.14 es werden alle ungspeicherten zurückgeliefert und hier dann entweder automatisch gespeichert (autosave) oder eine meldung ausgegeben
                    bool autosave_create = DataAdapter.Instance.DataCache.Profile.Option_GetValueBoolean(CSEnumOptions.autosave_create);
                    bool autosave_edit = DataAdapter.Instance.DataCache.Profile.Option_GetValueBoolean(CSEnumOptions.autosave_edit);
                    List<IDocumentOrFolder> autosaves = new List<IDocumentOrFolder>();
                    List<IDocumentOrFolder> notautosaves = new List<IDocumentOrFolder>();
                    foreach (IDocumentOrFolder obj in objectstosaveinstruct)
                    {
                        if ((obj.isNotCreated && autosave_create) || (obj.isEdited && autosave_edit))
                            autosaves.Add(obj);
                        else
                            notautosaves.Add(obj);
                    }
                    if (autosaves.Count == 0 && notautosaves.Count == 0)
                        ret = true;
                    else
                    {
                        if (autosaves.Count > 0)
                        {
                            bool valid = true;
                            string validatefailid = "";
                            try
                            {
                                foreach (IDocumentOrFolder obj in autosaves)
                                {
                                    validatefailid = obj.objectId;
                                    if (valid) valid = obj.ValidateData();
                                    if (!valid)
                                        break;
                                }
                            }
                            catch (Exception e) { Log.Log.Error(e); }
                            if (valid)
                                ret = true;
                        }
                        if (notautosaves.Count > 0)
                        {
                            ret = true;
                        }
                        if (!ret && displaywarning)
                        {
                            // meldung ausgeben
                            DisplayWarnMessage(LocalizationMapper.Instance["msg_warn_unsaved"]);
                        }
                    }
                }
            }
            return ret;
        }

        public bool _CanSetSelectedObject()
        {
            return (DataAdapter.Instance.DataCache.ApplicationFullyInit);
        }

        //private void SimulateScanSources()
        //{
        //    CSInformation[] infowithoutcommand = new CSInformation[4];
        //    CSInformation lcinfo1 = new CSInformation();
        //    lcinfo1.informationid = CSEnumInformationId.LC_Command;
        //    lcinfo1.informationvalue = "scan";
        //    lcinfo1.informationidSpecified = true;

        //    CSInformation lcinfo2 = new CSInformation();
        //    lcinfo2.informationid = CSEnumInformationId.LC_Param;
        //    lcinfo2.informationvalue = "scan";
        //    lcinfo2.informationidSpecified = true;

        //    CSInformation infoscan1 = new CSInformation();
        //    infoscan1.informationid = CSEnumInformationId.LC_Modul_Scan;
        //    infoscan1.informationvalue = "kofax";
        //    infoscan1.informationidSpecified = true;

        //    CSInformation infoscan2 = new CSInformation();
        //    infoscan2.informationid = CSEnumInformationId.LC_Modul_ScanParam;
        //    //infoscan2.informationvalue = "Scanvorlage Test1 verlängert|Scanvorlage Test2 verlängert|Scanvorlage Test3 verlängert|Scanvorlage Test4 verlängert|Scanvorlage Test5 verlängert|Scanvorlage Test6|Scanvorlage Test7";
        //    infoscan2.informationvalue = "Test1|Test2|Test3|Test4";
        //    infoscan2.informationidSpecified = true;

        //    infowithoutcommand[0] = lcinfo1;
        //    infowithoutcommand[1] = lcinfo2;
        //    infowithoutcommand[2] = infoscan1;
        //    infowithoutcommand[3] = infoscan2;

        //    DataAdapter.Instance.DataCache.ApplyData(infowithoutcommand);
        //    DataAdapter.Instance.InformObservers(CSEnumInformationId.LC_Modul_ScanParam.ToString());
        //}

        public void SetLastSelectedWorkspace(CSEnumProfileWorkspace newWorkspace)
        {
            if (LastSelectedWorkspace != newWorkspace)
            {
                LastSelectedWorkspace = newWorkspace;
                DataAdapter.Instance.InformObservers();
            }
        }

        public void SetLastSelectedRootnode(CSRootNode rootNode)
        {
            if (LastSelectedRootNode != rootNode)
            {
                LastSelectedRootNode = rootNode;
                DataAdapter.Instance.InformObservers();
            }
        }

        #endregion SetSelectedObject (cmd_SetSelectedObject)

        // ==================================================================

        #region SetSelectedObjects

        public void SetSelectedObjects(List<string> objectids)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanSetSelectedObjects)
                {
                    List<IDocumentOrFolder> objlist = new List<IDocumentOrFolder>();
                    foreach (string id in objectids)
                    {
                        IDocumentOrFolder tmp = DataAdapter.Instance.DataCache.Objects(Workspace).GetObjectById(id);
                        objlist.Add(tmp);
                        // TS 01.04.15 raus aus der schleife
                        //DataAdapter.Instance.DataCache.Objects(Workspace).SetSelectedObjects(objlist);
                        //DataAdapter.Instance.InformObservers();
                    }
                    // TS 01.04.15 raus aus der schleife
                    DataAdapter.Instance.DataCache.Objects(Workspace).SetSelectedObjects(objlist);
                    DataAdapter.Instance.InformObservers();
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        //public void SetSelectedObject(string objectid) { SetSelectedObject(objectid, 0, null); }
        //private void SetSelectedObject(object parameter) { SetSelectedObject((string)parameter); }
        // ------------------------------------------------------------------
        public bool CanSetSelectedObjects { get { return _CanSetSelectedObjects(); } }

        private bool _CanSetSelectedObjects()
        {
            return (DataAdapter.Instance.DataCache.ApplicationFullyInit);
        }

        #endregion SetSelectedObjects

        // ==================================================================

        #region StartSlideShow
        public void StartSlideShow()
        {
            // Clear Slideshow-List and start a new slideshow
            DataAdapter.Instance.DataCache.Objects(this.Workspace).ObjectIdList_SlideShow.Clear();
            DataAdapter.Instance.DataCache.Info.Viewer_Slideshow_Workspace = this.Workspace;
            DataAdapter.Instance.DataCache.Info.Viewer_SlideshowPaging = false;
            DataAdapter.Instance.DataCache.Info.Viewer_StartSlideshow = true;
        }
        #endregion

        // ==================================================================

        #region ShowRelation (cmd_ShowRelation, cmd_ShowRelationSelf)

        public DelegateCommand cmd_ShowRelation { get { return new DelegateCommand(ShowRelation, CanShowRelation); } }
        // TS 01.06.17
        //public DelegateCommand cmd_ShowRelationSelf { get { return new DelegateCommand(ShowRelationSelf, CanShowRelation); } }
        public DelegateCommand cmd_ShowRelationSelf { get { return new DelegateCommand(ShowRelationSelf, CanShowRelationSelf); } }
        public DelegateCommand cmd_NavigateRelation { get { return new DelegateCommand(NavigateRelation, CanNavigateRelation); } }
        public DelegateCommand cmd_NavigateRelationSelf { get { return new DelegateCommand(NavigateRelationSelf, CanNavigateRelation); } }

        public void ShowRelation(object relobject)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanShowRelation(relobject))
                    _ShowRelation(relobject, false);
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public void ShowRelationSelf(object relobject)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanShowRelation(relobject))
                    _ShowRelation(relobject, true);
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public void NavigateRelation(object relobject)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                // TS 29.05.17 canshowrelation liefert true wenn das objekt anzeigbar ist und z.b. nicht im externen html
                if (CanShowRelation(relobject))
                    _NavigateRelation(relobject, false, this.Workspace);
                // cannavigaterelation liefert true auch wenn es extern html anzeigbar ist
                else if (CanNavigateRelation(relobject))
                {
                    CSRCPushPullUIRouting routing = new CSRCPushPullUIRouting();
                    CSEnumRCPushUIListeners?[] receiver = new CSEnumRCPushUIListeners?[1];
                    receiver[0] = CSEnumRCPushUIListeners.HTML_CLIENT_DESKTOP;
                    routing.sendto_componentorvirtualids = receiver;

                    string repid = ((DocOrFolderRelationship)relobject).target_repid;
                    string mandantid = "0";
                    if (repid.Contains("_"))
                    {
                        mandantid = repid.Substring(repid.IndexOf("_") + 1);
                        repid = repid.Substring(0, repid.IndexOf("_"));
                    }
                    string objectid = ((DocOrFolderRelationship)relobject).target_objectid;
                    string extcommand = Statics.ExternalCommandHelper.CreateExternalCommand(repid, mandantid, objectid);

                    // If the HTML-Desktop is already open, send a push-command, else open the desktop with advanced params
                    if (Init.ViewManager.IsHTMLDesktopOpen())
                    {                        
                        List<string> objectidlist = new List<string>();
                        objectidlist.Add(extcommand);

                        DataAdapter.Instance.DataProvider.RCPush_UI(routing, CSEnumRCPushCommands.display, objectidlist, null, DataAdapter.Instance.DataCache.Rights.UserPrincipal);
                    }else
                    {
                        // Open the desktop-tool
                        IDocumentOrFolder selection = DataAdapter.Instance.DataCache.Objects(this.Workspace).Object_Selected;
                        string link = "2C_" + selection.RepositoryId + "_" + selection.objectId;
                        Toggle_ShowDesk("display=" + extcommand + "&ecmselection=" + link);
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public void NavigateRelationSelf(object relobject)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanShowRelationSelf(relobject))
                {
                    _NavigateRelation(relobject, true, this.Workspace);
                }
                else if (CanNavigateRelation(relobject))
                {
                    CSRCPushPullUIRouting routing = new CSRCPushPullUIRouting();
                    CSEnumRCPushUIListeners?[] receiver = new CSEnumRCPushUIListeners?[1];
                    receiver[0] = CSEnumRCPushUIListeners.HTML_CLIENT_DESKTOP;
                    routing.sendto_componentorvirtualids = receiver;

                    string repid = ((DocOrFolderRelationship)relobject).source_repid;
                    string mandantid = "0";
                    if (repid.Contains("_"))
                    {
                        mandantid = repid.Substring(repid.IndexOf("_") + 1);
                        repid = repid.Substring(0, repid.IndexOf("_"));
                    }
                    string objectid = ((DocOrFolderRelationship)relobject).source_objectid;
                    string extcommand = Statics.ExternalCommandHelper.CreateExternalCommand(repid, mandantid, objectid);

                    // If the HTML-Desktop is already open, send a push-command, else open the desktop with advanced params
                    if (Init.ViewManager.IsHTMLDesktopOpen())
                    {
                        List<string> objectidlist = new List<string>();
                        objectidlist.Add(extcommand);

                        DataAdapter.Instance.DataProvider.RCPush_UI(routing, CSEnumRCPushCommands.display, objectidlist, null, DataAdapter.Instance.DataCache.Rights.UserPrincipal);
                    }
                    else
                    {
                        // Open the desktop-tool
                        IDocumentOrFolder selection = DataAdapter.Instance.DataCache.Objects(this.Workspace).Object_Selected;
                        string link = "2C_" + selection.RepositoryId + "_" + selection.objectId;
                        Toggle_ShowDesk("display=" + extcommand + "&ecmselection=" + link);
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public void NavigateRelationFromOrigin(object relObject, bool isSelfRelation, CSEnumProfileWorkspace originWorkspace)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                // TS 01.06.17
                //if (CanShowRelation(relObject))
                //    _NavigateRelation(relObject, isSelfRelation, originWorkspace);
                if (isSelfRelation && CanShowRelationSelf(relObject))
                    _NavigateRelation(relObject, isSelfRelation, originWorkspace);
                else if (!isSelfRelation && CanShowRelation(relObject))
                    _NavigateRelation(relObject, isSelfRelation, originWorkspace);
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        private void _ShowRelation(object relobject, bool selfrelation)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                DocOrFolderRelationship relation = (DocOrFolderRelationship)relobject;

                if (!selfrelation)
                {
                    // TS 23.07.13 damit zuerst das objekt selektiert wird und dann die relation gelesen
                    // sonst wird während des relation lesens z.b. durch den listview auf das standard-objekt gesetzt
                    // TS 21.08.13 nur wenn nicht bereits gewählt
                    if (!DataAdapter.Instance.DataCache.Objects(Workspace).Object_Selected.objectId.Equals(relation.source_objectid))
                        SetSelectedObject(relation.source_objectid);
                }

                string ws = this.Workspace.ToString();
                if (!ws.EndsWith("_related"))
                    ws = ws + "_related";
                CSEnumProfileWorkspace wsrelated = EnumHelper.GetCSEnumFromValue<CSEnumProfileWorkspace>(ws);
                if (DataAdapter.Instance.DataCache.WorkspacesUsed.Contains(wsrelated))
                {
                    // related workspace clearen
                    // TS 11.02.14 der komplette clear prozess macht zu viele probleme durch die update-ui calls, besser schmal löschen da sowieso nur in related workspace
                    //DataAdapter.Instance.Processing(wsrelated).ClearCache();
                    DataAdapter.Instance.DataCache.Objects(wsrelated).ClearCache(false);

                    // TS 25.08.16 test only
                    //CallbackAction callback = new CallbackAction(_ShowRelationDone, relation.source_objectid, selfrelation, relobject);
                    CallbackAction callback = null;
                    if (selfrelation)
                        DataAdapter.Instance.Processing(wsrelated).QuerySingleObjectById(relation.source_repid, relation.source_objectid, false, true, wsrelated, callback);
                    else
                        DataAdapter.Instance.Processing(wsrelated).QuerySingleObjectById(relation.target_repid, relation.target_objectid, relation.target_isDocument, true, wsrelated, callback);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        /// <summary>
        /// Callback-Method for ShowRelation-Command
        /// Fires a following "NavigateRelation" if it's pending for special workspaces
        /// </summary>
        /// <param name="sourceObjectID"></param>
        /// <param name="isSelfRelation"></param>
        /// <param name="relObj"></param>
        // TS 25.08.16 test only
        //private void _ShowRelationDone(object sourceObjectID, object isSelfRelation, object relObj)
        //{
        //    try
        //    {
        //        string srcObjId = (string)sourceObjectID;
        //        bool isSelfRel = (bool)isSelfRelation;

        //        if (srcObjId.Length > 0)
        //        {
        //            IDocumentOrFolder srcObj = DataAdapter.Instance.DataCache.Objects(this.Workspace).GetObjectById(srcObjId);
        //            if (srcObj != null && srcObj.objectId.Length > 0 && srcObj.isNavigateRelationPending)
        //            {
        //                NavigateRelationFromOrigin(relObj, isSelfRel, this.Workspace);
        //                srcObj.isNavigateRelationPending = false;
        //            }
        //        }
        //    }
        //    catch (Exception ex) { if (ex.Message.Length > 0) { Log.Log.Debug(ex.Message); } }
        //}

        /// <summary>
        /// ermittelt den targetworkspace für eine relation
        /// und schickt eine meldung dorthin mittels visualobserver
        /// um das objekt der relation dort in ihrem kontext anzeigen zu lassen
        /// </summary>
        /// <param name="relobject"></param>
        /// <param name="selfrelations"></param>
        private void _NavigateRelation(object relobject, bool selfrelations, CSEnumProfileWorkspace originWorkspace)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                DocOrFolderRelationship relation = (DocOrFolderRelationship)relobject;

                // TS 16.08.16
                bool handled = false;
                bool wsset = false;

                // ================================================================
                string repositoryid = "";
                string relatedobjectid = "";
                IDocumentOrFolder sourceObject = DataAdapter.Instance.DataCache.Objects(originWorkspace).GetObjectById(relation.source_objectid);

                char splitter = ("|".ToCharArray())[0];

                if (selfrelations)
                {
                    repositoryid = relation.source_repid;
                    relatedobjectid = relation.source_objectid;
                }
                else
                {
                    repositoryid = relation.target_repid;
                    relatedobjectid = relation.target_objectid;
                }

                CSEnumProfileWorkspace targetworkspace = CSEnumProfileWorkspace.workspace_undefined;
                CSEnumProfileWorkspace detailworkspace = CSEnumProfileWorkspace.workspace_undefined;
                string reptryidshort = repositoryid;
                if (reptryidshort != null && reptryidshort.Length > 0 && reptryidshort.Contains("_"))
                    reptryidshort = reptryidshort.Substring(0, reptryidshort.IndexOf("_"));
                int reptryid = int.Parse(reptryidshort);
                switch (reptryid)
                {
                    // TS 08.08.16 KIL nicht mehr vorhanden
                    //case Constants.REPOSITORY_KIL:
                    //    targetworkspace = CSEnumProfileWorkspace.workspace_default;
                    //    detailworkspace = CSEnumProfileWorkspace.workspace_default;
                    //    break;
                    case Constants.REPOSITORY_MAIL:
                        targetworkspace = CSEnumProfileWorkspace.workspace_edesktop;
                        detailworkspace = CSEnumProfileWorkspace.workspace_mail;
                        break;

                    case Constants.REPOSITORY_POST:
                        targetworkspace = CSEnumProfileWorkspace.workspace_edesktop;
                        detailworkspace = CSEnumProfileWorkspace.workspace_post;
                        break;

                    case Constants.REPOSITORY_INTERNAL:
                        string name = relation.cmisname;
                        if (name != null && name.Length > 0)
                        {
                            string[] tokens = name.Split(splitter);
                            if (tokens.Length == 3)
                                name = tokens[1];
                        }
                        targetworkspace = CSEnumProfileWorkspace.workspace_edesktop;
                        if (name.Equals(CSEnumInternalObjectType.Ablage.ToString()))
                            detailworkspace = CSEnumProfileWorkspace.workspace_ablage;
                        else if (name.Equals(CSEnumInternalObjectType.Aufgabe.ToString()))
                            detailworkspace = CSEnumProfileWorkspace.workspace_aufgabe;
                        else if (name.Equals(CSEnumInternalObjectType.AufgabePublic.ToString()))
                            detailworkspace = CSEnumProfileWorkspace.workspace_aufgabepublic;
                        else if (name.Equals(CSEnumInternalObjectType.Stapel.ToString()))
                            detailworkspace = CSEnumProfileWorkspace.workspace_stapel;
                        else if (name.Equals(CSEnumInternalObjectType.Termin.ToString()))
                            detailworkspace = CSEnumProfileWorkspace.workspace_termin;
                        else if (name.Equals(CSEnumInternalObjectType.Workflow.ToString()))
                            detailworkspace = CSEnumProfileWorkspace.workspace_workflow;                        
                        else if (name.Contains(CSEnumInternalObjectType.Favorit.ToString()))
                        {
                            // TS 16.08.16
                            targetworkspace = CSEnumProfileWorkspace.workspace_favoriten;
                            detailworkspace = CSEnumProfileWorkspace.workspace_favoriten;
                        }
                        else if (name.Contains(CSEnumInternalObjectType.Link.ToString()))
                        {
                            // TS 16.08.16
                            IDocumentOrFolder relatedobject = DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_favoriten_related).GetObjectById(relatedobjectid);

                            if (relatedobject != null && relatedobject.objectId.Length > 0)
                            {
                                string internal01 = relatedobject[CSEnumCmisProperties.D_INTERNAL_01];
                                string internal02 = relatedobject[CSEnumCmisProperties.D_INTERNAL_02];
                                // TS 16.08.16
                                if (internal01 != null && internal01.Equals(CSEnumInternalObjectType.Favorit.ToString()))
                                {
                                    targetworkspace = CSEnumProfileWorkspace.workspace_favoriten;
                                    detailworkspace = CSEnumProfileWorkspace.workspace_favoriten;
                                    wsset = true;
                                }
                                else if (internal01 != null && internal01.Equals(CSEnumInternalObjectType.Link.ToString()))
                                {
                                    if (internal02 != null && internal02.Equals(CSEnumInternalObjectType.LinkQuery.ToString()))
                                    {
                                        // =============================
                                        // deserialize
                                        this.ClearCache();
                                        string deserialized = ObjectSerializing.ReadLongStringFromPropertyAllValues(relatedobject, CSEnumCmisProperties.D_INTERNAL_10);
                                        List<CSQueryProperties> res = ObjectSerializing.DeSerializeQueryProperties(deserialized, true);
                                        DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_default).Query_LastQueryValues = res;
                                        DataAdapter.Instance.Processing(CSEnumProfileWorkspace.workspace_default).QuerySetLast();
                                        DataAdapter.Instance.Processing(CSEnumProfileWorkspace.workspace_default).Refresh();
                                        handled = true;
                                    }
                                    else if (internal02 != null && internal02.Equals(CSEnumInternalObjectType.LinkFS.ToString()))
                                    {
                                        string attribs = relatedobject[CSEnumCmisProperties.D_INTERNAL_10];
                                        // call LC
                                        if (attribs != null && attribs.Length > 0)
                                        {
                                            //attribs = "@" + attribs;
                                            this.LC_FileSelect(attribs);
                                        }
                                        handled = true;
                                    }
                                    else if (internal02 != null && internal02.Equals(CSEnumInternalObjectType.LinkURL.ToString()))
                                    {
                                        string attribs = relatedobject[CSEnumCmisProperties.D_INTERNAL_10];
                                        DisplayExternal(attribs);
                                        handled = true;
                                    }
                                }
                            }
                            // TS 16.08.16
                            if (!wsset)
                            {
                                targetworkspace = CSEnumProfileWorkspace.workspace_favoriten_related;
                                detailworkspace = CSEnumProfileWorkspace.workspace_favoriten_related;
                            }
                        }
                        break;

                    default:
                        targetworkspace = CSEnumProfileWorkspace.workspace_default;
                        detailworkspace = CSEnumProfileWorkspace.workspace_default;
                        break;
                }
                // TS 16.08.16
                if (!handled)
                {
                    // TS 19.08.16 rootnodewechsel wenn anderes repository (zunächst noch abgeklemmt)
                    bool nodechange = false;
                    if (targetworkspace == CSEnumProfileWorkspace.workspace_default)
                    {
                        // if (!repositoryid.Equals(DataAdapter.Instance.DataCache.RootNodes.Node_Selected.repositoryid))
                        if (!repositoryid.Equals(DataAdapter.Instance.DataCache.Rights.ChoosenRepository.repositoryId))
                        {
                            nodechange = true;
                            CSRootNode rootnode = DataAdapter.Instance.DataCache.RootNodes.GetNodeById(repositoryid);
                            //object relobject, bool selfrelations, CSEnumProfileWorkspace originWorkspace
                            AppManager.RestartRelationObject = (DocOrFolderRelationship)relobject;
                            AppManager.RestartRelationIsSelf = selfrelations;
                            AppManager.RestartRelationWorkspace = originWorkspace;
                            AppManager.ChooseRootNode(rootnode, false);
                        }
                    }

                    if (!nodechange)
                    {
                        if (!detailworkspace.ToString().Equals(this.Workspace.ToString()))
                        {
                            // mitteilung versenden
                            DataAdapter.Instance.NavigateUI(targetworkspace, detailworkspace);
                            // relation im zielworkspace anzeigen
                            DataAdapter.Instance.Processing(detailworkspace).NavigateRelationFromOrigin(relobject, selfrelations, originWorkspace);
                        }
                        else
                        {
                            IDocumentOrFolder relatedobject = DataAdapter.Instance.DataCache.Objects(this.Workspace).GetObjectById(relatedobjectid);
                            if (relatedobject != null && relatedobject.objectId.Length > 0)
                            {
                                SetSelectedObject(relatedobjectid);
                            }
                            else
                            {
                                // TS 25.08.16 test only
                                QuerySingleObjectById(repositoryid, relatedobjectid, false, true, this.Workspace, null);
                                //// Supress the further processing on special workspaces, do it later
                                //if (sourceObject.isNavigateRelationPending || !Statics.Constants.GetWorkspacesWithoutShowRelationLink().Contains(originWorkspace))
                                //{
                                //    QuerySingleObjectById(repositoryid, relatedobjectid, false, true, this.Workspace, null);
                                //}
                                //else
                                //{
                                //    if (sourceObject != null && sourceObject.objectId.Length > 0)
                                //    {
                                //        sourceObject.isNavigateRelationPending = true;
                                //    }
                                //}
                            }
                        }
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public bool CanShowRelation(object relobject) { return CanShowRelation(relobject, true); }

        public bool CanShowRelationSelf(object relobject) { return CanShowRelation(relobject, true); }

        public bool CanNavigateRelation(object relobject) { return CanShowRelation(relobject, false); }
        

        private bool CanShowRelation(object relobject, bool checkforexternalhtml)
        {
            bool ret = false;
            if (DataAdapter.Instance.DataCache.ApplicationFullyInit)
            {
                string ws = this.Workspace.ToString();
                if (!ws.EndsWith("_related"))
                    ws = ws + "_related";
                CSEnumProfileWorkspace wsrelated = EnumHelper.GetCSEnumFromValue<CSEnumProfileWorkspace>(ws);
                ret = DataAdapter.Instance.DataCache.WorkspacesUsed.Contains(wsrelated);
                if (ret) ret = relobject != null;
                if (ret) ret = relobject.GetType().Name.Equals("DocOrFolderRelationship");
                // TS 29.05.17 wenn extern (HTML Desktop) dann soll es nicht angezeigt werden
                // die unterscheidung wird verwendet bei Show_Relation und Navigate_Relation
                // CanShow_Relation liefert false wenn HTML Desktop verwendet wird und es eine Relation auf den Posteingang oder eine Aufgabe ist
                // Can_Navigate_Relation liefert true wenn HTML Desktop verwendet wird und es eine Relation auf den Posteingang ist und sendet dann ein display an HTML
                if (ret && ViewManager.IsHTMLDesktop && checkforexternalhtml)
                {
                    string repid = ((DocOrFolderRelationship)relobject).target_repid;
                    DocOrFolderRelationship relationShip = (DocOrFolderRelationship)relobject;
                    if (repid != null && repid.Length > 0)
                    {
                        if (repid.Contains("_"))
                            repid = repid.Substring(0, repid.IndexOf("_"));
                        if (int.Parse(repid) == Statics.Constants.REPOSITORY_POST)
                        {
                            ret = false;
                        }else if(relationShip.cmisname.Contains("Aufgabe"))
                        {
                            ret = false;
                        }
                    }
                }
            }
            return ret;
        }

        // ------------------------------------------------------------------

        #endregion ShowRelation (cmd_ShowRelation, cmd_ShowRelationSelf)

        // ==================================================================

        #region Follow_Favorite (cmd_Follow_Favorite)
        public DelegateCommand cmd_Follow_Favorite { get { return new DelegateCommand(Follow_Favorite, CanFollow_Favorite); } }
        public void Follow_Favorite()
        {
            if (_CanFollow_Favorite())
            {
                IDocumentOrFolder selFav = DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_favoriten).Object_Selected;
                foreach(DocOrFolderRelationship rel in selFav.Relationships)
                {
                    DataAdapter.Instance.Processing(CSEnumProfileWorkspace.workspace_default).NavigateRelation(rel);
                    break;
                }
            }
        }

        public bool CanFollow_Favorite()
        {
            return _CanFollow_Favorite();
        }

        private bool _CanFollow_Favorite()
        {
            return DataAdapter.Instance.DataCache.ApplicationFullyInit && this.Workspace == CSEnumProfileWorkspace.workspace_favoriten
                && DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_favoriten).Object_Selected != null
                && DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_favoriten).Object_Selected.objectId.Length > 0;
        }
        #endregion

        // ==================================================================

        #region JumpToOrigin (cmd_JumpToOrigin)
        public DelegateCommand cmd_JumpToOrigin { get { return new DelegateCommand(JumpToOrigin, CanJumpToOrigin); } }
        public void JumpToOrigin()
        {
            try
            {
                if (_CanJumpToOrigin())
                {
                    IDocumentOrFolder selObj = DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_gesamt).Object_Selected;
                    if (selObj != null && selObj.objectId.Length > 0)
                    {
                        // Catch the object-type for navigating
                        string extender = selObj.isAufgabe ? CSEnumInternalObjectType.Aufgabe.ToString() : selObj.isStapel ? CSEnumInternalObjectType.Stapel.ToString() : selObj.isTermin ? CSEnumInternalObjectType.Termin.ToString() : selObj[CSEnumCmisProperties.BPM_DOC_TYPE] != null ? CSEnumInternalObjectType.Workflow.ToString() : "";
                        string cmisName = selObj.objectId + "|" + extender + "|" + selObj.objectId;

                        // Create a dummy-relationship to use default-methods
                        DocOrFolderRelationship dummy_rel = new DocOrFolderRelationship(selObj.objectId, selObj.RepositoryId, selObj.objectId, "", selObj.RepositoryId, selObj.objectId, "", cmisName, "", "");
                        NavigateRelationSelf(dummy_rel);
                    }
                }
            }
            catch (System.Exception ex) { if (ex.Message.Length > 0) { Log.Log.Debug(ex.Message); } }
        }

        public bool CanJumpToOrigin()
        {
            return _CanJumpToOrigin();
        }

        private bool _CanJumpToOrigin()
        {
            return DataAdapter.Instance.DataCache.ApplicationFullyInit && this.Workspace == CSEnumProfileWorkspace.workspace_gesamt
                && DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_gesamt).Object_Selected != null
                && DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_gesamt).Object_Selected.objectId.Length > 0;
        }
        #endregion

        // ==================================================================

        #region ShowPendingTasks (cmd_ShowPendingTasks)
        public DelegateCommand cmd_ShowPendingTasks { get { return new DelegateCommand(ShowPendingTasks, _CanShowPendingTasks); } }
        public void ShowPendingTasks()
        {
            if (_CanShowPendingTasks())
            {
                View.Dialogs.dlgPendingTasks dlgPendingTask = new View.Dialogs.dlgPendingTasks();
                DialogHandler.Show_Dialog(dlgPendingTask);
            }
        }

        private bool _CanShowPendingTasks()
        {
            return DataAdapter.Instance.DataCache.ApplicationFullyInit;
        }
        #endregion

        // ==================================================================

        #region ShowAnnos (cmd_ShowAnnos)

        public DelegateCommand cmd_ShowAnnos { get { return new DelegateCommand(ShowAnnos, _CanShowAnnos); } }

        public void ShowAnnos()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanShowAnnos)
                {
                    CSOption showannos = DataAdapter.Instance.DataCache.Profile.Option_GetOption(CSEnumOptions.viewershowannos);
                    showannos.value = true.ToString();
                    DataAdapter.Instance.InformObservers();
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public bool CanShowAnnos { get { return _CanShowAnnos(); } }

        private bool _CanShowAnnos()
        {
            return DataAdapter.Instance.DataCache.ApplicationFullyInit
                && DataAdapter.Instance.DataCache.Objects(this.Workspace).Document_Selected.objectId.Length > 0
                && !DataAdapter.Instance.DataCache.Profile.Option_GetValueBoolean(CSEnumOptions.viewershowannos);
        }

        #endregion ShowAnnos (cmd_ShowAnnos)

        // ==================================================================

        #region HideAnnos (cmd_HideAnnos)

        public DelegateCommand cmd_HideAnnos { get { return new DelegateCommand(HideAnnos, _CanHideAnnos); } }

        public void HideAnnos()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanHideAnnos)
                {
                    CSOption showannos = DataAdapter.Instance.DataCache.Profile.Option_GetOption(CSEnumOptions.viewershowannos);
                    showannos.value = false.ToString();
                    DataAdapter.Instance.InformObservers();
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public bool CanHideAnnos { get { return _CanHideAnnos(); } }

        private bool _CanHideAnnos()
        {
            return DataAdapter.Instance.DataCache.ApplicationFullyInit
                && DataAdapter.Instance.DataCache.Objects(this.Workspace).Document_Selected.objectId.Length > 0
                && DataAdapter.Instance.DataCache.Profile.Option_GetValueBoolean(CSEnumOptions.viewershowannos);
        }

        #endregion HideAnnos (cmd_HideAnnos)

        // ==================================================================

        #region ShowAddonMeta (cmd_ShowAddonMeta)

        public DelegateCommand cmd_ShowAddonMeta { get { return new DelegateCommand(ShowAddonMeta); } }

        public void ShowAddonMeta()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                //WindowOpen(CSEnumProfileComponentType.ADDONMETA, this.Workspace);
                CSProfileComponent profilenode = DataAdapter.Instance.DataCache.Profile.Profile_GetComponentFromLayout(CSEnumProfileComponentType.ADDONMETA, "");
                if (profilenode != null)
                {
                    if (profilenode.visible)
                        WindowClose(profilenode);
                    else
                        WindowOpen(CSEnumProfileComponentType.ADDONMETA, this.Workspace);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public bool CanShowAddonMeta { get { return _CanShowAddonMeta(); } }

        private bool _CanShowAddonMeta()
        {
            return DataAdapter.Instance.DataCache.ApplicationFullyInit;
        }

        #endregion ShowAddonMeta (cmd_ShowAddonMeta)

        // ==================================================================

        #region ShowAddonViewer (cmd_ShowAddonViewer)

        public DelegateCommand cmd_ShowAddonViewer { get { return new DelegateCommand(ShowAddonViewer); } }

        public void ShowAddonViewer()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                WindowOpen(CSEnumProfileComponentType.ADDONVIEWER, this.Workspace);
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public bool CanShowAddonViewer { get { return _CanShowAddonViewer(); } }

        private bool _CanShowAddonViewer()
        {
            return DataAdapter.Instance.DataCache.ApplicationFullyInit;
        }

        #endregion ShowAddonViewer (cmd_ShowAddonViewer)

        // ==================================================================

        #region ShowAddonTree (cmd_ShowAddonTree)

        public DelegateCommand cmd_ShowAddonTree { get { return new DelegateCommand(ShowAddonTree); } }

        public void ShowAddonTree()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                WindowOpen(CSEnumProfileComponentType.ADDONTREE, this.Workspace);
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public bool CanShowAddonTree { get { return _CanShowAddonTree(); } }

        private bool _CanShowAddonTree()
        {
            return DataAdapter.Instance.DataCache.ApplicationFullyInit;
        }

        #endregion ShowAddonTree (cmd_ShowAddonTree)

        // ==================================================================

        #region ShowAddonMatchlist (cmd_ShowAddonMatchlist)

        public DelegateCommand cmd_ShowAddonMatchlist { get { return new DelegateCommand(ShowAddonMatchlist); } }

        public void ShowAddonMatchlist()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                WindowOpen(CSEnumProfileComponentType.ADDONMATCHLIST, CSEnumProfileWorkspace.workspace_searchoverall);
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public bool CanShowAddonMatchlist { get { return _CanShowAddonMatchlist(); } }

        private bool _CanShowAddonMatchlist()
        {
            return DataAdapter.Instance.DataCache.ApplicationFullyInit;
        }

        #endregion ShowAddonMatchlist (cmd_ShowAddonMatchlist)

        // ==================================================================

        #region ShowAddonThumbs (cmd_ShowAddonThumbs)

        public DelegateCommand cmd_ShowAddonThumbs { get { return new DelegateCommand(ShowAddonThumbs); } }

        public void ShowAddonThumbs()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                WindowOpen(CSEnumProfileComponentType.ADDONTHUMBS, this.Workspace);
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public bool CanShowAddonThumbs { get { return _CanShowAddonThumbs(); } }

        private bool _CanShowAddonThumbs()
        {
            return DataAdapter.Instance.DataCache.ApplicationFullyInit;
        }

        #endregion ShowAddonThumbs (cmd_ShowAddonThumbs)

        // ==================================================================

        #region ShowAktenplan (cmd_ShowAktenplan)

        public DelegateCommand cmd_ShowAktenplan { get { return new DelegateCommand(ShowAktenplan, _CanShowAktenplan); } }

        public void ShowAktenplan()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                // TS 25.11.15
                //WindowOpen(CSEnumProfileComponentType.AKTENPLAN, this.Workspace);
                WindowOpen(CSEnumProfileComponentType.AKTENPLAN, CSEnumProfileWorkspace.workspace_aktenplan);
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public bool CanShowAktenplan { get { return _CanShowAktenplan(); } }

        private bool _CanShowAktenplan()
        {
            bool ret = DataAdapter.Instance.DataCache.ApplicationFullyInit;
            if (ret)
            {
                CSProfileComponent comp = DataAdapter.Instance.DataCache.Profile.Profile_GetFirstComponentOfType(CSEnumProfileComponentType.AKTENPLAN);
                ret = comp != null && !comp.visible;
                // TS 24.09.15 nur wenn knoten vorhanden sind
                if (ret) ret = DataAdapter.Instance.DataCache.RootNodes.NodeList.Count > 0 && !DataAdapter.Instance.DataCache.RootNodes.Node_Selected.autogenerated;
            }
            return ret;
        }

        #endregion ShowAktenplan (cmd_ShowAktenplan)

        // ==================================================================

        #region ShowProtocol (cmd_ShowProtocol)

        public DelegateCommand cmd_ShowProtocol { get { return new DelegateCommand(ShowProtocol, _CanShowProtocol); } }

        public void ShowProtocol()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanShowProtocol)
                {
                    // Call the webservice for getting the protocol entries of the given object
                    // TODO: Parametrize the maxitem-count :)
                    IDocumentOrFolder selObj = DataAdapter.Instance.DataCache.Objects(Workspace).Folder_Selected;
                    CallbackAction callback = new CallbackAction(ShowProtocol_Done, selObj.objectId);
                    DataAdapter.Instance.DataProvider.GetProtocol(selObj, 50, DataAdapter.Instance.DataCache.Rights.UserPrincipal, callback);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        private void ShowProtocol_Done(string objID)
        {
            // TS 26.06.17 abfangen wenn nicht gelesen werden konnte
            if (DataAdapter.Instance.DataCache.ResponseStatus.success)
            {
                IDocumentOrFolder objSelected = DataAdapter.Instance.DataCache.Objects(this.Workspace).GetObjectById(objID);
                if (objSelected != null && objSelected.objectId.Length > 0)
                {
                    View.Dialogs.dlgObjectsProtocol protDialog = new View.Dialogs.dlgObjectsProtocol(objSelected);
                    DialogHandler.Show_Dialog(protDialog);
                }
            }
            else
            {
                DisplayMessage(LocalizationMapper.Instance["msg_insufficient_rights"]);
            }
        }

        public bool CanShowProtocol { get { return _CanShowProtocol(); } }

        private bool _CanShowProtocol()
        {
            bool ret = DataAdapter.Instance.DataCache.ApplicationFullyInit;
            if (ret)
            {
                //// TODO Prüfe Rechte hier oder so
                // ret = DataAdapter.Instance.DataCache.Repository(this.Workspace).HasPermission(DataCache_Repository.RepositoryPermissions.PERMISSION_PROTOKOLL_5);

                IDocumentOrFolder selObj = DataAdapter.Instance.DataCache.Objects(this.Workspace).Folder_Selected;
                if(selObj.ACL != null)
                {
                    ret = selObj.hasReadProtocolPermission();
                }
                else
                {
                    ret = false;
                }
            }
            return ret;
        }    

        #endregion ShowProtocol (cmd_ShowProtocol)

        // ==================================================================


        #region ShowCopyRelations (cmd_ShowCopyRelations)

        public DelegateCommand cmd_ShowCopyRelations { get { return new DelegateCommand(ShowCopyRelations, _CanShowCopyRelations); } }

        public void ShowCopyRelations()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanShowCopyRelations)
                {
                    IDocumentOrFolder selObj = DataAdapter.Instance.DataCache.Objects(this.Workspace).Object_Selected;
                    if(selObj != null && selObj.objectId.Length > 0)
                    {
                        CallbackAction callback = new CallbackAction(_Query_Done);
                        List<CSQueryProperties> propList = new List<CSQueryProperties>();
                        ClearCache();

                        // Query for copies
                        if (selObj.hasCopyRelations)
                        {
                            // Prepare QueryProperties                            
                            foreach (DocOrFolderRelationship rel in selObj.getCopyRelations)
                            {
                                propList.AddRange(QueryCreatePropertiesList(rel.source_repid, rel.source_objectid, rel.source_objectid.StartsWith("D")));
                            }
                        }
                        else if(selObj.isCopyRelation)
                        // Query for origin
                        {
                            propList.AddRange(QueryCreatePropertiesList(selObj.getCopyRelationOrigin.target_repid, selObj.getCopyRelationOrigin.target_objectid, selObj.getCopyRelationOrigin.target_objectid.StartsWith("D")));
                        }

                        // Shot the query
                        if (propList.Count > 0)
                        {
                            int querylistsize = int.Parse(DataAdapter.Instance.DataCache.Profile.Option_GetValue(CSEnumOptions.querylistsize));
                            DataAdapter.Instance.DataProvider.Query(propList, false, true, false, false, querylistsize, CSEnumProfileWorkspace.workspace_searchoverall, DataAdapter.Instance.DataCache.Rights.UserPrincipal, callback);
                        }
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public bool CanShowCopyRelations { get { return _CanShowCopyRelations(); } }

        private bool _CanShowCopyRelations()
        {
            bool ret = false;
            try
            {
                IDocumentOrFolder selObj = DataAdapter.Instance.DataCache.Objects(this.Workspace).Object_Selected;

                if (selObj != null && selObj.objectId.Length > 0
                    && DataAdapter.Instance.DataCache.ApplicationFullyInit
                    && (selObj.isCopyRelation || selObj.hasCopyRelations))
                {
                    ret = true;
                }
            }
            catch (Exception e) { Log.Log.Error(e); }

            return ret;
        }

        #endregion ShowCopyRelations (cmd_ShowCopyRelations)

        // ==================================================================


        #region ShowDeleted (cmd_ShowDeleted)

        public DelegateCommand cmd_ShowDeleted { get { return new DelegateCommand(ShowDeleted, _CanShowDeleted); } }

        public void ShowDeleted()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanShowDeleted)
                {
                    SL2C_Client.View.Dialogs.dlgShowDeleted child = new SL2C_Client.View.Dialogs.dlgShowDeleted(CSEnumProfileWorkspace.workspace_deleted);
                    DialogHandler.Show_Dialog(child);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public bool CanShowDeleted { get { return _CanShowDeleted(); } }

        private bool _CanShowDeleted()
        {
            return DataAdapter.Instance.DataCache.ApplicationFullyInit;
        }

        #endregion ShowDeleted (cmd_ShowDeleted)

        // ==================================================================

        #region ShowLastEdited (cmd_ShowLastEdited)

        public DelegateCommand cmd_ShowLastEdited { get { return new DelegateCommand(ShowLastEdited, _CanShowLastEdited); } }

        public void ShowLastEdited()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanShowLastEdited)
                {
                    SL2C_Client.View.Dialogs.dlgShowLastEdited child = new SL2C_Client.View.Dialogs.dlgShowLastEdited(CSEnumProfileWorkspace.workspace_lastedited);
                    DialogHandler.Show_Dialog(child);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public bool CanShowLastEdited { get { return _CanShowLastEdited(); } }

        private bool _CanShowLastEdited()
        {
            return DataAdapter.Instance.DataCache.ApplicationFullyInit;
        }

        #endregion ShowLastEdited (cmd_ShowLastEdited)

        // ==================================================================

        // ==================================================================

        #region ShowLinksFromObject (cmd_ShowLinksFromObject)

        public DelegateCommand cmd_ShowLinksFromObject { get { return new DelegateCommand(ShowLinksFromObject_PreRunner, _CanShowLinksFromObject); } }

        private void ShowLinksFromObject_PreRunner(object selObj)
        {
            IDocumentOrFolder selobject = (IDocumentOrFolder)selObj;
            DataAdapter.Instance.DataProvider.GetObject(selobject.RepositoryId, selobject.objectId, this.Workspace, DataAdapter.Instance.DataCache.Rights.UserPrincipal, new CallbackAction(ShowLinksFromObject, selobject));
        }

        public void ShowLinksFromObject(object selObj)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                IDocumentOrFolder selItem = (IDocumentOrFolder)selObj;
                SetSelectedObject(selItem.objectId);
                if (CanShowLinksFromObject)
                {
                    SL2C_Client.View.Dialogs.dlgShowLinksFromObject child = new SL2C_Client.View.Dialogs.dlgShowLinksFromObject(CSEnumProfileWorkspace.workspace_links, (IDocumentOrFolder)selObj);
                    DialogHandler.Show_Dialog(child);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public bool CanShowLinksFromObject { get { return _CanShowLinksFromObject(); } }

        private bool _CanShowLinksFromObject()
        {
            return DataAdapter.Instance.DataCache.ApplicationFullyInit;
        }

        #endregion ShowLinksFromObject (cmd_ShowLinksFromObject)

        // ==================================================================

        #region ShowAdressLinksFromObject (cmd_ShowAdressLinksFromObject)

        public DelegateCommand cmd_ShowAdressLinksFromObject { get { return new DelegateCommand(ShowAdressLinksFromObject_PreRunner, _CanShowAdressLinksFromObject); } }

        private void ShowAdressLinksFromObject_PreRunner(object selObj)
        {
            IDocumentOrFolder selobject = (IDocumentOrFolder)selObj;
            DataAdapter.Instance.DataProvider.GetObject(selobject.RepositoryId, selobject.objectId, this.Workspace, DataAdapter.Instance.DataCache.Rights.UserPrincipal, new CallbackAction(ShowAdressLinksFromObject, selobject));
        }

        public void ShowAdressLinksFromObject(object selObj)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                IDocumentOrFolder selItem = (IDocumentOrFolder)selObj;
                SetSelectedObject(selItem.objectId);
                if (CanShowAdressLinksFromObject)
                {
                    SL2C_Client.View.Dialogs.dlgShowAdressLinksFromObject child = new SL2C_Client.View.Dialogs.dlgShowAdressLinksFromObject(CSEnumProfileWorkspace.workspace_links, (IDocumentOrFolder)selObj);
                    DialogHandler.Show_Dialog(child);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public bool CanShowAdressLinksFromObject { get { return _CanShowAdressLinksFromObject(); } }

        private bool _CanShowAdressLinksFromObject()
        {
            return DataAdapter.Instance.DataCache.ApplicationFullyInit;
        }

        #endregion ShowAdressLinksFromObject (cmd_ShowAdressLinksFromObject)

        // ==================================================================

        #region ShowTaskLinksFromObject (cmd_ShowTaskLinksFromObject)

        public DelegateCommand cmd_ShowTaskLinksFromObject { get { return new DelegateCommand(ShowTaskLinksFromObject_PreRunner, _CanShowTaskLinksFromObject); } }

        private void ShowTaskLinksFromObject_PreRunner(object selObj)
        {
            IDocumentOrFolder selobject = (IDocumentOrFolder)selObj;
            DataAdapter.Instance.DataProvider.GetObject(selobject.RepositoryId, selobject.objectId, this.Workspace, DataAdapter.Instance.DataCache.Rights.UserPrincipal, new CallbackAction(ShowTaskLinksFromObject, selobject));
        }

        public void ShowTaskLinksFromObject(object selObj)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                IDocumentOrFolder selItem = (IDocumentOrFolder)selObj;
                SetSelectedObject(selItem.objectId);
                if (CanShowTaskLinksFromObject)
                {
                    SL2C_Client.View.Dialogs.dlgShowTaskLinksFromObject child = new SL2C_Client.View.Dialogs.dlgShowTaskLinksFromObject(CSEnumProfileWorkspace.workspace_links, (IDocumentOrFolder)selObj);
                    DialogHandler.Show_Dialog(child);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public bool CanShowTaskLinksFromObject { get { return _CanShowTaskLinksFromObject(); } }

        private bool _CanShowTaskLinksFromObject()
        {
            return DataAdapter.Instance.DataCache.ApplicationFullyInit;
        }

        #endregion ShowTaskLinksFromObject (cmd_ShowTaskLinksFromObject)

        // ==================================================================

        #region ShowHelp (cmd_ShowHelp)

        public DelegateCommand cmd_ShowHelp { get { return new DelegateCommand(ShowHelp, _CanShowHelp); } }

        public void ShowHelp_Hotkey() { ShowHelp(); }

        public void ShowHelp()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanShowHelp)
                {
                    string url = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.URL_Documentation];
                    if (url == null || url.Length == 0)
                    {
                        List<CSInformation> infoasked = new List<CSInformation>();

                        // url von doku holen
                        CSInformation info = new CSInformation();
                        info.informationid = CSEnumInformationId.URL_Documentation;
                        info.informationidSpecified = true;
                        infoasked.Add(info);

                        // versionsinfos gleich mit holen
                        info = new CSInformation();
                        info.informationid = CSEnumInformationId.VersionCServer;
                        info.informationidSpecified = true;
                        infoasked.Add(info);

                        info = new CSInformation();
                        info.informationid = CSEnumInformationId.VersionDMSServer;
                        info.informationidSpecified = true;
                        infoasked.Add(info);

                        DataAdapter.Instance.DataProvider.GetInformationList(infoasked, DataAdapter.Instance.DataCache.Rights.UserPrincipal, new CallbackAction(ShowHelp_Done));
                    }
                    else
                        ShowHelp_Done();
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public bool CanShowHelp { get { return _CanShowHelp(); } }

        private bool _CanShowHelp()
        {
            return DataAdapter.Instance.DataCache.ApplicationFullyInit;
        }

        private void ShowHelp_Done()
        {
            string url = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.URL_Documentation];
            if (url != null && url.Length > 0)
            {
                //System.Windows.Browser.HtmlWindow newWindow;
                //newWindow = HtmlPage.Window.Navigate(new Uri(downloadname, UriKind.RelativeOrAbsolute), "_blank", "");
                // TS 08.07.15
                //HtmlPage.Window.Navigate(new Uri(url, UriKind.RelativeOrAbsolute), "_blank", "");
                DisplayExternal(url);
            }
        }

        #endregion ShowHelp (cmd_ShowHelp)

        // ==================================================================

        #region ShowInfo (cmd_ShowInfo)

        public DelegateCommand cmd_ShowInfo { get { return new DelegateCommand(ShowInfo, _CanShowInfo); } }

        public void ShowInfo()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanShowInfo)
                {
                    string url = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.VersionCServer];
                    if (url == null || url.Length == 0)
                    {
                        List<CSInformation> infoasked = new List<CSInformation>();

                        // versionsinfos holen
                        CSInformation info = new CSInformation();
                        info.informationid = CSEnumInformationId.VersionCServer;
                        info.informationidSpecified = true;
                        infoasked.Add(info);

                        info = new CSInformation();
                        info.informationid = CSEnumInformationId.VersionDMSServer;
                        info.informationidSpecified = true;
                        infoasked.Add(info);

                        // url von doku gleich mit holen
                        info = new CSInformation();
                        info.informationid = CSEnumInformationId.URL_Documentation;
                        info.informationidSpecified = true;
                        infoasked.Add(info);

                        DataAdapter.Instance.DataProvider.GetInformationList(infoasked, DataAdapter.Instance.DataCache.Rights.UserPrincipal, new CallbackAction(ShowInfo_Done));
                    }
                    else
                        ShowInfo_Done();

                    //CSRCPushPullUIRouting routing = new CSRCPushPullUIRouting();
                    //string[] receiver = new string[1];
                    //receiver[0] = Statics.Constants.UILISTENER_VIRTUALUNIQUEID_HTML_CLIENT;
                    //routing.sendto_componentorvirtualids = receiver;

                    //string repid = ((DocOrFolderRelationship)relobject).target_repid;
                    //string mandantid = "0";
                    //if (repid.Contains("_"))
                    //{
                    //    mandantid = repid.Substring(repid.IndexOf("_") + 1);
                    //    repid = repid.Substring(0, repid.IndexOf("_"));
                    //}
                    //string objectid = ((DocOrFolderRelationship)relobject).target_objectid;
                    //string extcommand = Statics.ExternalCommandHelper.CreateExternalCommand(repid, mandantid, objectid);

                    //List<string> objectidlist = new List<string>();
                    //objectidlist.Add(extcommand);

                    //DataAdapter.Instance.DataProvider.RCPush_UI(routing, CSEnumRCPushCommands.display, objectidlist, null, DataAdapter.Instance.DataCache.Rights.UserPrincipal);

                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public bool CanShowInfo { get { return _CanShowInfo(); } }

        private bool _CanShowInfo()
        {
            return DataAdapter.Instance.DataCache.ApplicationFullyInit;
        }

        private void ShowInfo_Done()
        {
            if (DataAdapter.Instance.DataCache.ResponseStatus.success)
            {
                string cversion = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.VersionCServer];
                string dversion = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.VersionDMSServer];
                SL2C_Client.View.Dialogs.dlgVersion child = new SL2C_Client.View.Dialogs.dlgVersion(cversion, dversion);
                DialogHandler.Show_Dialog(child);
            }
        }

        #endregion ShowInfo (cmd_ShowInfo)

        // ==================================================================

        #region ShowSubPageDesktop (cmd_ShowSubPageDesktop)

        public DelegateCommand cmd_ShowSubPageDesktop { get { return new DelegateCommand(ShowSubPageDesktop, _CanShowSubPageDesktop); } }

        public void ShowSubPageDesktop()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanShowSubPageDesktop && LocalStorage.IsAllDataCached)
                {
                    // Create the subpages' link                  
                    string pageURL = DataAdapter.Instance.DataProvider.GetProviderAttribs(true, true, true, true, false); ;
                    string extensionPage = Constants.SUBPAGE_COMMUNICATION_PARAM_SUBPAGE + "=" + Constants.EnumSubPageTypes.SubPage_Desktop;
                    string extensionWorkspace = Constants.SUBPAGE_COMMUNICATION_PARAM_WORKSPACE + "=" + EnumHelper.GetValueFromCSEnum(CSEnumProfileWorkspace.workspace_edesktop);
                    pageURL += "?" + extensionPage + "&" + extensionWorkspace;

                    // Open the browser-window
                    HtmlPage.Window.Navigate(new Uri(pageURL), "_blank", "toolbar=no,location=no,status=no,menubar=no,resizable=yes");
                    
                    if(isMasterApplication)
                    {
                        Toggle_ShowDefault();
                        DataAdapter.Instance.NavigateUI(CSEnumProfileWorkspace.workspace_default, CSEnumProfileWorkspace.workspace_default);
                    }
                }else
                {
                    if(showMessageBox(LocalizationMapper.Instance["msg_browsercache_not_used"], MessageBoxButton.OK) == MessageBoxResult.OK)
                    {
                        SaveCache(true, true, true, true, true, true, new CallbackAction(ShowSubPageDesktop));
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public bool CanShowSubPageDesktop { get { return _CanShowSubPageDesktop(); } }

        private bool _CanShowSubPageDesktop()
        {
            return DataAdapter.Instance.DataCache.ApplicationFullyInit && isMasterApplication;
        }

        #endregion ShowSubPageDesktop (cmd_ShowSubPageDesktop)

        // ==================================================================

        #region ShowSubPageViewer (cmd_ShowSubPageViewer)

        public DelegateCommand cmd_ShowSubPageViewer { get { return new DelegateCommand(ShowSubPageViewer, _CanShowSubPageViewer); } }

        public void ShowSubPageViewer()
        {            
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanShowSubPageViewer && LocalStorage.IsAllDataCached)
                {
                    // Create the subpages' link
                    string pageURL = DataAdapter.Instance.DataProvider.GetProviderAttribs(true, true, true, true, false); ;
                    string extensionPage = Constants.SUBPAGE_COMMUNICATION_PARAM_SUBPAGE + "=" + Constants.EnumSubPageTypes.SubPage_Viewer;
                    string extensionWorkspace = Constants.SUBPAGE_COMMUNICATION_PARAM_WORKSPACE + "=" + EnumHelper.GetValueFromCSEnum(CSEnumProfileWorkspace.workspace_undefined);
                    pageURL += "?" + extensionPage + "&" + extensionWorkspace;

                    // Open the browser-window
                    HtmlPage.Window.Navigate(new Uri(pageURL), "_blank", "toolbar=no,location=no,status=no,menubar=no,resizable=yes");
                }
                else
                {
                    if (showMessageBox(LocalizationMapper.Instance["msg_browsercache_not_used"], MessageBoxButton.OK) == MessageBoxResult.OK)
                    {
                        SaveCache(true, true, true, true, true, true, new CallbackAction(ShowSubPageViewer));                        
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public bool CanShowSubPageViewer { get { return _CanShowSubPageViewer(); } }

        private bool _CanShowSubPageViewer()
        {
            return DataAdapter.Instance.DataCache.ApplicationFullyInit && isMasterApplication;
        }

        #endregion ShowSubPageViewer (cmd_ShowSubPageViewer)     

        // ==================================================================

        #region ShowSubPageTreeview (cmd_ShowSubPageTreeview)

        public DelegateCommand cmd_ShowSubPageTreeview { get { return new DelegateCommand(ShowSubPageTreeview, _CanShowSubPageTreeview); } }

        public void ShowSubPageTreeview()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanShowSubPageTreeview && LocalStorage.IsAllDataCached)
                {
                    // Create the subpages' link
                    string pageURL = DataAdapter.Instance.DataProvider.GetProviderAttribs(true, true, true, true, false); ;
                    string extensionPage = Constants.SUBPAGE_COMMUNICATION_PARAM_SUBPAGE + "=" + Constants.EnumSubPageTypes.SubPage_Tree;
                    string extensionWorkspace = Constants.SUBPAGE_COMMUNICATION_PARAM_WORKSPACE + "=" + EnumHelper.GetValueFromCSEnum(CSEnumProfileWorkspace.workspace_default);
                    pageURL += "?" + extensionPage + "&" + extensionWorkspace;

                    // Open the browser-window
                    HtmlPage.Window.Navigate(new Uri(pageURL), "_blank", "toolbar=no,location=no,status=no,menubar=no,resizable=yes");
                }
                else
                {
                    if (showMessageBox(LocalizationMapper.Instance["msg_browsercache_not_used"], MessageBoxButton.OK) == MessageBoxResult.OK)
                    {
                        SaveCache(true, true, true, true, true, true, new CallbackAction(ShowSubPageTreeview));
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public bool CanShowSubPageTreeview { get { return _CanShowSubPageTreeview(); } }

        private bool _CanShowSubPageTreeview()
        {
            return DataAdapter.Instance.DataCache.ApplicationFullyInit && isMasterApplication;
        }

        #endregion ShowSubPageTreeview (cmd_ShowSubPageTreeview)

        // ==================================================================

        #region ShowOptions (cmd_ShowOptions)

        public DelegateCommand cmd_ShowOptions { get { return new DelegateCommand(ShowOptions, _CanShowOptions); } }

        public void ShowOptions_Hotkey()
        {
            ShowOptions(null);
        }

        public void ShowOptions(object profilenode)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanShowOptions)
                {
                    if (profilenode == null) profilenode = DataAdapter.Instance.DataCache.Profile.UserProfile;

                    SL2C_Client.View.Dialogs.Options.dlgOptions child = new SL2C_Client.View.Dialogs.Options.dlgOptions(profilenode);
                    DialogHandler.Show_Dialog(child);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public bool CanShowOptions { get { return _CanShowOptions(); } }

        private bool _CanShowOptions()
        {
            return DataAdapter.Instance.DataCache.ApplicationFullyInit;
        }

        #endregion ShowOptions (cmd_ShowOptions)

        // ==================================================================

        #region ShowProperties (cmd_ShowProperties)

        public DelegateCommand cmd_ShowProperties { get { return new DelegateCommand(ShowProperties, _CanShowProperties); } }

        public void ShowProperties_Hotkey() { ShowProperties(); }
        public void ShowProperties()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanShowProperties)
                {
                    IDocumentOrFolder dummy = DataAdapter.Instance.DataCache.Objects(Workspace).Object_Selected;
                    // TS 28.09.17
                    if (dummy.objectId.Length == 0)
                    {
                        dummy = DataAdapter.Instance.DataCache.Objects(Workspace).Root;
                    }

                    if (dummy.ACL == null && dummy.canGetACL)
                        DataAdapter.Instance.DataProvider.GetACL(dummy.CMISObject, DataAdapter.Instance.DataCache.Rights.UserPrincipal, new CallbackAction(ShowPropertiesGetACLDone));
                    else
                        ShowPropertiesGetACLDone();
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // ------------------------------------------------------------------
        public bool CanShowProperties { get { return _CanShowProperties(); } }

        private bool _CanShowProperties()
        {
            // TS 28.09.17
            //return DataAdapter.Instance.DataCache.ApplicationFullyInit && DataAdapter.Instance.DataCache.Objects(Workspace).Object_Selected.objectId.Length > 0;
            return DataAdapter.Instance.DataCache.ApplicationFullyInit;
        }

        // ------------------------------------------------------------------
        private void ShowPropertiesGetACLDone()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (DataAdapter.Instance.DataCache.ResponseStatus.success)
                {
                    // TS 28.09.17
                    //SL2C_Client.View.Dialogs.Properties.dlgProperties child = new SL2C_Client.View.Dialogs.Properties.dlgProperties(DataAdapter.Instance.DataCache.Objects(Workspace).Object_Selected);
                    IDocumentOrFolder dummy = DataAdapter.Instance.DataCache.Objects(Workspace).Object_Selected;
                    if (dummy.objectId.Length == 0)
                    {
                        dummy = DataAdapter.Instance.DataCache.Objects(Workspace).Root;
                    }
                    SL2C_Client.View.Dialogs.Properties.dlgProperties child = new SL2C_Client.View.Dialogs.Properties.dlgProperties(dummy);
                    DialogHandler.Show_Dialog(child);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        #endregion ShowProperties (cmd_ShowProperties)

        // ==================================================================

        #region WindowClose (cmd_WindowClose)

        public DelegateCommand cmd_WindowClose { get { return new DelegateCommand(WindowClose, _CanWindowClose); } }

        public void WindowClose(object profilenode)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanWindowClose)
                {
                    if (profilenode != null)
                    {
                        CSProfileComponent givencomponent = (CSProfileComponent)profilenode;

                        // TS 08.02.16 clipboard sonderbehandlung
                        if (givencomponent != null && givencomponent.type == CSEnumProfileComponentType.CLIPBOARD)
                        {
                            if (CanClearClipboard) this.ClearClipboard();
                        }
                        else
                        {
                            // TS 25.01.13 umbau
                            //foreach (CSProfileComponent component in DataAdapter.Instance.DataCache.Profile.ProfileLayout.components)
                            //{
                            CSProfileComponent component = DataAdapter.Instance.DataCache.Profile.Profile_GetComponentFromLayout(givencomponent.type, givencomponent.id);

                            if (component.id.Equals(givencomponent.id))
                            {
                                // wenn subcomponente dann den parent holen
                                CSProfileComponent parent = DataAdapter.Instance.DataCache.Profile.Profile_GetTopComponentOfSubComponent(component);
                                if (parent != null)
                                    component = parent;

                                // update profile info
                                component.visible = false;
                                // hide or close window in backcanvas
                                if (component.keepalive)
                                {
                                    ViewManager.HideComponent(component);
                                    DataAdapter.Instance.InformObservers();
                                }
                                else
                                {
                                    //ViewManager.UnloadComponent(component.id);
                                    ViewManager.UnloadComponent(component);
                                    // die prüfung ob gelöscht werden darf findet dort statt
                                    DataAdapter.Instance.DataCache.Profile.Profile_RemoveComponent(component);
                                    DataAdapter.Instance.InformObservers();
                                }
                                //break;
                            }
                            //}
                        }
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public bool CanWindowClose { get { return _CanWindowClose(); } }

        private bool _CanWindowClose()
        {
            return DataAdapter.Instance.DataCache.ApplicationFullyInit;
            //bool ret = DataAdapter.Instance.DataCache.ApplicationFullyInit;
            //if (ret && this.Workspace.ToString().Equals(CSEnumProfileWorkspace.workspace_clipboard.ToString()))
            //    ret = DataAdapter.Instance.DataCache.Objects(this.Workspace).ObjectList.Count < 2;
            //return ret;
        }

        #endregion WindowClose (cmd_WindowClose)

        // ==================================================================

        #region WindowOpen (cmd_WindowOpen)

        public DelegateCommand cmd_WindowOpen { get { return new DelegateCommand(WindowOpen, _CanWindowOpen); } }

        public void WindowOpen(object componenttype)
        {
            WindowOpen(componenttype, null);
        }

        public void WindowOpen(object componenttype, object enumprofileworkspace)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (_CanWindowOpen(componenttype))
                {
                    CSEnumProfileComponentType profilecomponenttype = CSEnumProfileComponentType.UNDEFINED;
                    CSProfileComponent component = null;
                    if (componenttype.GetType().Name.Equals("CSEnumProfileComponentType"))
                        profilecomponenttype = (CSEnumProfileComponentType)componenttype;
                    else if (componenttype.GetType().Name.Equals("String"))
                    {
                        // TS 20.06.16
                        string ctype = (string)componenttype;
                        if (ctype.StartsWith("ID="))
                        {
                            ctype = ctype.Substring(3);
                            component = DataAdapter.Instance.DataCache.Profile.Profile_GetComponentFromLayout(ctype);
                            profilecomponenttype = component.type;
                        }
                        else
                        {
                            profilecomponenttype = (CSEnumProfileComponentType)Enum.Parse(typeof(CSEnumProfileComponentType), (string)componenttype, true);
                        }
                    }
                    // TS 20.06.16 dialog extra prüfen, der rest komplett in die klammer
                    if (component != null && profilecomponenttype == CSEnumProfileComponentType.DIALOGEXTERNAL)
                    {
                        // DIALOGEXTERNAL
                        View.Dialogs.dlgExternal ext = new View.Dialogs.dlgExternal(component);
                        //ext.ComponentId = component.id;
                        DialogHandler.Show_Dialog(ext);
                    }
                    else
                    {
                        if (profilecomponenttype != CSEnumProfileComponentType.UNDEFINED)
                        {
                            // TS 15.11.13 für addontypen zunächst geschlossene komponenten vom basetype suchen
                            //component = DataAdapter.Instance.DataCache.Profile.Profile_GetInvisibleComponentOfType(profilecomponenttype, enumprofileworkspace);
                            if (ProfileComponentTypeAttribs.IsAddonType(profilecomponenttype))
                            {
                                CSEnumProfileComponentType basetype = ProfileComponentTypeAttribs.GetBaseType(profilecomponenttype);
                                component = DataAdapter.Instance.DataCache.Profile.Profile_GetInvisibleComponentOfType(basetype, enumprofileworkspace);
                            }
                            // TS 15.11.13
                            if (component == null)
                                component = DataAdapter.Instance.DataCache.Profile.Profile_GetInvisibleComponentOfType(profilecomponenttype, enumprofileworkspace);
                        }

                        if (component != null)
                        {
                            component.visible = true;
                            if (component.keepalive)
                            {
                                ViewManager.ShowComponent(component);
                            }
                            else
                            {
                                ViewManager.LoadComponent(component);
                            }
                        }
                        else if (profilecomponenttype != CSEnumProfileComponentType.UNDEFINED)
                        {
                            // TS 30.08.13
                            //component = DataAdapter.Instance.DataCache.Profile.Profile_CreateComponent(profilecomponenttype);
                            component = DataAdapter.Instance.DataCache.Profile.Profile_CreateComponent(profilecomponenttype, enumprofileworkspace);
                            if (component != null)
                            {
                                component.visible = true;
                                ViewManager.LoadComponent(component);
                            }
                        }

                        // TS 15.11.13 hier nach unten geholt
                        // TS 27.08.13 wenn toplevel window dann andere toplevelwindows ggf. schliessen
                        if (component != null
                            && DataAdapter.Instance.DataCache.Profile.Option_GetValueBoolean(component.options, CSEnumOptions.istoplevelwindow)
                            && DataAdapter.Instance.DataCache.Profile.Option_GetValueBoolean(component.options, CSEnumOptions.forceclosetoplevel))
                        {
                            List<CSProfileComponent> othertoplevelwindows = DataAdapter.Instance.DataCache.Profile.Profile_GetComponentsHavingOption(CSEnumOptions.istoplevelwindow);
                            if (othertoplevelwindows != null && othertoplevelwindows.Count > 0)
                            {
                                foreach (CSProfileComponent comp in othertoplevelwindows)
                                {
                                    if (!comp.id.Equals(component.id) && comp.type != CSEnumProfileComponentType.CLIPBOARD)
                                    {
                                        bool istoplevel = DataAdapter.Instance.DataCache.Profile.Option_GetValueBoolean(comp.options, CSEnumOptions.istoplevelwindow);
                                        bool autoclose = DataAdapter.Instance.DataCache.Profile.Option_GetValueBoolean(comp.options, CSEnumOptions.autocloseontoplevel);
                                        if (istoplevel && comp.visible && autoclose)
                                        {
                                            WindowClose(comp);
                                        }
                                    }
                                }
                            }
                        }

                        if (component != null)
                            DataAdapter.Instance.InformObservers();

                        // TS 15.11.13 wenn addonviewer dann gleiche größe wie aktueller viewer und gleiche zoomstufe
                        //if (component != null && component.type.ToString().Equals(CSEnumProfileComponentType.ADDONVIEWER.ToString()))
                        //    _AdjustViewerZoom(component);
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        private bool _CanWindowOpen(object componenttype)
        {
            bool ret = false;
            if (DataAdapter.Instance.DataCache.ApplicationFullyInit && componenttype != null)
            {
                CSEnumProfileComponentType profilecomponenttype;
                CSProfileComponent component = null;
                if (componenttype.GetType().Name.Equals("CSEnumProfileComponentType"))
                {
                    profilecomponenttype = (CSEnumProfileComponentType)componenttype;
                    component = DataAdapter.Instance.DataCache.Profile.Profile_GetVisibleComponentOfType(profilecomponenttype);
                }
                else if (componenttype.GetType().Name.Equals("String"))
                {
                    // TS 20.06.16
                    string ctype = (string)componenttype;
                    if (ctype.StartsWith("ID="))
                    {
                        ctype = ctype.Substring(3);
                        component = DataAdapter.Instance.DataCache.Profile.Profile_GetComponentFromLayout(ctype);
                    }
                    else
                    {
                        profilecomponenttype = (CSEnumProfileComponentType)Enum.Parse(typeof(CSEnumProfileComponentType), (string)componenttype, true);
                        component = DataAdapter.Instance.DataCache.Profile.Profile_GetVisibleComponentOfType(profilecomponenttype);
                    }
                }
                if (component != null)
                {
                    ret = (!component.keepalive);
                }
                else
                    ret = true;
            }
            return ret;
        }

        //// TS 15.11.13
        //private void _AdjustViewerZoom(CSProfileComponent component)
        //{
        //    // TS 15.11.13 groesse anpassen auf letzte sichtbare komponente des gleichen typs
        //    if (component != null && component.type.ToString().Equals(CSEnumProfileComponentType.ADDONVIEWER.ToString()))
        //    {
        //        CSProfileComponent template = DataAdapter.Instance.DataCache.Profile.Profile_GetLastComponentOfType(CSEnumProfileComponentType.VIEWER, true);
        //        if (template != null)
        //        {
        //            component.height = template.height;
        //            component.width = template.width;

        //            IVisualObserver observercomp = DataAdapter.Instance.FindVisualObserverForProfileComponent(component);
        //            IVisualObserver observertemp = DataAdapter.Instance.FindVisualObserverForProfileComponent(template);
        //            if (observercomp != null && observertemp != null)
        //            {
        //                IPageBindingAdapter pbacomp = observercomp.GetPageBindingAdapter();
        //                IPageBindingAdapter pbatemp = observertemp.GetPageBindingAdapter();
        //                if (pbacomp != null && pbatemp != null)
        //                {
        //                    IPageProcessing pprcomp = pbacomp.PageProcessing;
        //                    IPageProcessing pprtemp = pbatemp.PageProcessing;
        //                    if (pprcomp != null && pprtemp != null)
        //                    {
        //                        ViewerProcessing vprcomp = (ViewerProcessing)pprcomp;
        //                        ViewerProcessing vprtemp = (ViewerProcessing)pprtemp;
        //                        vprcomp.MSI.ViewportWidth = vprtemp.MSI.ViewportWidth;
        //                    }
        //                }
        //            }
        //        }
        //    }
        //}

        #endregion WindowOpen (cmd_WindowOpen)

        // ==================================================================

        #region WindowDock (cmd_WindowDock)

        public DelegateCommand cmd_WindowDock { get { return new DelegateCommand(WindowDock, _CanWindowDock); } }

        public void WindowDock(object profilenode)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (_CanWindowDock(profilenode))
                {
                    CSProfileComponent givencomponent = (CSProfileComponent)profilenode;
                    if (ViewManager.ComponentCanDock(givencomponent) && !ViewManager.IsComponentDocked(givencomponent) && ViewManager.DockComponent(givencomponent))
                    {
                        givencomponent.docked = true;
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        //public bool CanWindowDock { get { return _CanWindowDock(); } }
        private bool _CanWindowDock(object profilenode)
        {
            if (DataAdapter.Instance.DataCache.ApplicationFullyInit && profilenode != null)
            {
                return true;
            }
            else
                return false;
        }

        #endregion WindowDock (cmd_WindowDock)

        // ==================================================================

        #region WindowUndock (cmd_WindowUndock)

        public DelegateCommand cmd_WindowUndock { get { return new DelegateCommand(WindowUndock, _CanWindowUndock); } }

        public void WindowUndock(object profilenode)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (_CanWindowUndock(profilenode))
                {
                    CSProfileComponent givencomponent = (CSProfileComponent)profilenode;
                    if (ViewManager.IsComponentDocked(givencomponent) && ViewManager.UndockComponent(givencomponent))
                    {
                        givencomponent.docked = false;
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }
      
        private bool _CanWindowUndock(object profilenode)
        {
            if (DataAdapter.Instance.DataCache.ApplicationFullyInit && profilenode != null)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        #endregion WindowUndock (cmd_WindowUndock)

        // ==================================================================

        #region Toggle_ShowDefault (cmd_Toggle_ShowDefault)

        // TS 08.03.13 speziell für die togglebuttons ohne can_ property damit die buttons nicht disabled werden
        // stattdessen wird die ischecked property an das can_ gebunden über eine zu erstellende Property im BindingAdapter
        //
        // !!! WICHTIG: IM BindingAdapter (und in den Interfaces) muss die CAN_XYZ Property eingetragen werden !!!
        // !!! sonst funktioniert das nicht richtig
        //
        public DelegateCommand cmd_Toggle_ShowDefault { get { return new DelegateCommand(Toggle_ShowDefault); } }

        public void Toggle_ShowDefault()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (((MainPage)App.Current.RootVisual).IsSplittedView && !((MainPage)App.Current.RootVisual).IsViewDefault)
                    ((MainPage)App.Current.RootVisual).ShowDefault();

                CSProfileComponent profilenode_adr = DataAdapter.Instance.DataCache.Profile.Profile_GetComponentFromLayout(CSEnumProfileComponentType.ADRESSEN, Constants.PROFILE_COMPONENTID_ADRESSEN);
                CSProfileComponent profilenode_edp = DataAdapter.Instance.DataCache.Profile.Profile_GetComponentFromLayout(CSEnumProfileComponentType.EDESKTOP, Constants.PROFILE_COMPONENTID_EDESKTOP);

                if (profilenode_adr != null && !ViewManager.IsComponentDocked(profilenode_adr) && ViewManager.IsComponentVisible(profilenode_adr))
                    WindowClose(profilenode_adr);

                if (profilenode_edp != null && !ViewManager.IsComponentDocked(profilenode_edp) && ViewManager.IsComponentVisible(profilenode_edp))
                    WindowClose(profilenode_edp);

                // TS 15.06.15 adress workspace (und edesk) informieren damit der sich ggf. wieder aktiviert (UpdateUIPaused !!)
                DataAdapter.Instance.InformObservers();
                //DataAdapter.Instance.InformObservers(this.Workspace, "CanToggle_ShowDefault");
                //DataAdapter.Instance.InformObservers(this.Workspace, "CanToggle_ShowAdrvw");
                //DataAdapter.Instance.InformObservers(this.Workspace, "CanToggle_ShowDesk");
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public bool CanToggle_ShowDefault { get { return _CanToggle_ShowDefault(); } }

        private bool _CanToggle_ShowDefault()
        {
            bool ret = DataAdapter.Instance.DataCache.ApplicationFullyInit;
            return ret;
        }

        public bool IsEnabledToggle_ShowDefault { get { return DataAdapter.Instance.DataCache.ApplicationFullyInit; } }

        #endregion Toggle_ShowDefault (cmd_Toggle_ShowDefault)

        // ==================================================================

        #region Toggle_ShowAdrvw (cmd_Toggle_ShowAdrvw)

        // TS 08.03.13 speziell für die togglebuttons ohne can_ property damit die buttons nicht disabled werden
        // stattdessen wird die ischecked property an das can_ gebunden über eine zu erstellende Property im BindingAdapter
        //
        // !!! WICHTIG: IM BindingAdapter (und in den Interfaces) muss die CAN_XYZ Property eingetragen werden !!!
        // !!! sonst funktioniert das nicht richtig
        //
        public DelegateCommand cmd_Toggle_ShowAdrvw { get { return new DelegateCommand(Toggle_ShowAdrvw); } }

        public void Toggle_ShowAdrvw()
        {
            Toggle_ShowAdrvw("");
        }

        public void Toggle_ShowAdrvw(string urlparams)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                // Open HTML-Adress
                if (ViewManager.IsHTMLAdress)
                {
                    ViewManager.ShowHTMLAdress(urlparams);
                    DataAdapter.Instance.InformObservers();
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public bool CanToggle_ShowAdrvw { get { return _CanToggle_ShowAdrvw(); } }

        private bool _CanToggle_ShowAdrvw()
        {
            bool ret = DataAdapter.Instance.DataCache.ApplicationFullyInit;
            if (ret)
            {
                ret = false;
                CSProfileComponent profilenode = DataAdapter.Instance.DataCache.Profile.Profile_GetComponentFromLayout(CSEnumProfileComponentType.ADRESSEN, Constants.PROFILE_COMPONENTID_ADRESSEN);
                if (profilenode != null)
                {
                    if (!((MainPage)App.Current.RootVisual).IsSplittedView)
                        ret = profilenode.visible == false;
                    else
                        ret = !((MainPage)App.Current.RootVisual).IsViewAdressen;
                }
            }
            return ret;
        }

        public bool IsEnabledToggle_ShowAdrvw { get { return DataAdapter.Instance.DataCache.ApplicationFullyInit; } }

        public Visibility Toggle_ShowAdrvwAlert
        {
            get
            {
                if (CanToggle_ShowAdrvw && DataAdapter.Instance.DataCache.Objects(this.Workspace).Folder_Selected_Level(Constants.STRUCTLEVEL_07_AKTE).isAdressenAkte)
                    return Visibility.Visible;
                else
                    return Visibility.Collapsed;
            }
        }

        // TS 30.01.17
        public string Toggle_ShowAdrvwAlertText
        {
            get
            {
                if (CanToggle_ShowAdrvw && DataAdapter.Instance.DataCache.Objects(this.Workspace).Folder_Selected_Level(Constants.STRUCTLEVEL_07_AKTE).isAdressenAkte)
                    return "1";
                else
                    return "";
            }
        }

        #endregion Toggle_ShowAdrvw (cmd_Toggle_ShowAdrvw)

        // ==================================================================

        #region Toggle_ShowDesk (cmd_Toggle_ShowDesk)

        // TS 08.03.13 speziell für die togglebuttons ohne can_ property damit die buttons nicht disabled werden
        // stattdessen wird die ischecked property an das can_ gebunden über eine zu erstellende Property im BindingAdapter
        //
        // !!! WICHTIG: IM BindingAdapter (und in den Interfaces) muss die CAN_XYZ Property eingetragen werden !!!
        // !!! sonst funktioniert das nicht richtig
        //
        public DelegateCommand cmd_Toggle_ShowDesk { get { return new DelegateCommand(Toggle_ShowDesk); } }

        public void Toggle_ShowDesk()
        {
            IDocumentOrFolder selection = DataAdapter.Instance.DataCache.Objects(Workspace).Object_Selected;
            string link = "2C_" + selection.RepositoryId + "_" + selection.objectId;
            Toggle_ShowDesk("ecmselection=" + link);
        }

        public void Toggle_ShowDesk(string urlparams)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                //CSProfileComponent profilenode = DataAdapter.Instance.DataCache.Profile.Profile_GetComponentFromLayout(CSEnumProfileComponentType.EDESKTOP, Constants.PROFILE_COMPONENTID_EDESKTOP);
                //if (profilenode != null && CanToggle_ShowDesk)
                //{
                //    if (!((MainPage)App.Current.RootVisual).IsSplittedView || !ViewManager.IsComponentDocked(profilenode))
                //    {
                //        WindowOpen(profilenode.type);
                //    }
                //    else
                //    {
                //        ((MainPage)App.Current.RootVisual).ShowDesktop();
                //    }
                //    DataAdapter.Instance.DataCache.Info.Alert = false;
                //}

                // Open HTML-Desktop
                if (ViewManager.IsHTMLDesktop)
                {
                    ViewManager.ShowHTMLDesktop(urlparams);
                    DataAdapter.Instance.InformObservers();
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public bool CanToggle_ShowDesk { get { return _CanToggle_ShowDesk(); } }

        private bool _CanToggle_ShowDesk()
        {
            bool ret = DataAdapter.Instance.DataCache.ApplicationFullyInit;
            if (ret)
            {
                ret = false;
                CSProfileComponent profilenode = DataAdapter.Instance.DataCache.Profile.Profile_GetComponentFromLayout(CSEnumProfileComponentType.EDESKTOP, Constants.PROFILE_COMPONENTID_EDESKTOP);
                if (profilenode != null)
                {
                    if (!((MainPage)App.Current.RootVisual).IsSplittedView)
                        ret = profilenode.visible == false;
                    else
                        ret = !((MainPage)App.Current.RootVisual).IsViewDesktop;
                }
            }
            return ret;
        }

        public bool IsEnabledToggle_ShowDesk { get { return DataAdapter.Instance.DataCache.ApplicationFullyInit; } }

        public Visibility Toggle_ShowDeskAlert
        {
            get
            {
                // LF: Rausgenommen da das Alert-Symbol erstmal nicht erscheinen soll
                //if (CanToggle_ShowDesk && DataAdapter.Instance.DataCache.Info.Alert)
                //    return Visibility.Visible;
                //else
                    return Visibility.Collapsed;
            }
        }

        public string Toggle_ShowDeskAlertText
        {
            get
            {
                if (DataAdapter.Instance.DataCache.Info.CurrentUnreadCount > 0)
                    return DataAdapter.Instance.DataCache.Info.CurrentUnreadCount.ToString();
                else
                    return "";
            }
        }

        #endregion Toggle_ShowDesk (cmd_Toggle_ShowDesk)

        #region BPM_ToggleSideFrage

        /// <summary>
        /// Shows the sidewinder-frame for workflows or general side-data
        /// </summary>
        /// <param name="form"></param>
        /// <param name="div"></param>
        public void BPM_DisplaySideFrame(HtmlElement form, HtmlElement div, string url)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();

            try
            {
                if (form == null) form = HtmlPage.Document.GetElementById("form1");
                if (div == null) div = HtmlPage.Document.GetElementById("bpmSideDrop");

                HtmlElement bpmiFrame = HtmlPage.Document.GetElementById("bpmSideDropFrame");
                HtmlElement bpmContentDiv = HtmlPage.Document.GetElementById("bpmSideDropContent");

                if (form != null && div != null)
                {
                    string visible = div.GetStyleAttribute("display");
                    bool mustshow = (visible != null && visible.Equals("none"));

                    if (mustshow)
                    {
                        const string PX = "{0}px";

                        double width = (double)HtmlPage.Window.GetProperty("innerWidth");
                        double widthSideDrop = Statics.Constants.BPM_SIDEDROP_MIN_WIDTH;
                        double widthMainPage = width - widthSideDrop;
                        double divSideDropOffset = widthMainPage;

                        form.RemoveStyleAttribute("width");
                        form.SetStyleAttribute("width", widthMainPage + "px");
                        form.RemoveStyleAttribute("height");
                        form.SetStyleAttribute("height", "100%");

                        div.RemoveStyleAttribute("left");
                        div.RemoveStyleAttribute("top");
                        div.RemoveStyleAttribute("width");
                        div.RemoveStyleAttribute("height");
                        div.RemoveStyleAttribute("position");

                        div.SetStyleAttribute("position", "absolute");
                        div.SetStyleAttribute("left", string.Format(PX, divSideDropOffset));
                        div.SetStyleAttribute("top", "0");
                        div.SetStyleAttribute("width", (widthSideDrop - 1) + "px");
                        div.SetStyleAttribute("height", "100%");
                        bpmiFrame.SetStyleAttribute("width", "100%");
                        bpmiFrame.SetStyleAttribute("height", "100%");
                        bpmiFrame.SetStyleAttribute("border", "none");

                        div.RemoveStyleAttribute("display");
                        bpmiFrame.RemoveStyleAttribute("display");

                        //Adding additional style information
                        object colorDivider = App.Current.Resources["REF_BackColor_Darkest"];
                        if (colorDivider != null)
                        {
                            string hexDivider = ((System.Windows.Media.Color)colorDivider).ToString().Substring(3);

                            //Adding divider-styled border (for dragging)
                            bpmContentDiv.SetStyleAttribute("border-left", "3px solid #" + hexDivider);
                        }

                        //Toggling toolbar-layout
                        BPM_ToggleToolbarLayout(false);
                    }
                    if (url == null || url.Length == 0) url = "about:blank";
                    bpmiFrame.SetProperty("src", url);
                    if (DataAdapter.Instance.DataCache.Objects(Workspace).Folder_Selected != null && DataAdapter.Instance.DataCache.Objects(Workspace).Folder_Selected.objectId.Length > 0)
                    {
                        bpmiFrame.SetAttribute("currentWF", DataAdapter.Instance.DataCache.Objects(Workspace).Folder_Selected.objectId);
                    }
                    DataAdapter.Instance.InformObservers(CSEnumProfileWorkspace.workspace_default);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        private static bool _isBPM_Headerbar_Created = false;

        /// <summary>
        /// Toggles the layout of the toolbar html-toolbar extension.
        /// Provides a html-based extension of the silverlight-toolbar for a better view
        /// </summary>
        /// <param name="originalLayout"></param>        
        public void BPM_ToggleToolbarLayout(bool originalLayout)
        {
            try
            {
                //Visibility changeVisibility = originalLayout ? Visibility.Visible : Visibility.Collapsed;
                HtmlElement bpmSideDropHeader = HtmlPage.Document.GetElementById("bpmSideDropHeader");
                HtmlElement bpmSideDropHeaderImage = HtmlPage.Document.GetElementById("bpmSideDropHeaderImage");

                // Building the header-bar once
                if (!_isBPM_Headerbar_Created)
                {
                    // Coloring
                    object colorStart = App.Current.Resources["REF_BackColor_Dark"];
                    object colorEnd = App.Current.Resources["REF_BackColor_Lightest"];
                    if (colorStart != null && colorEnd != null)
                    {
                        string hexStart = ((System.Windows.Media.Color)colorStart).ToString().Substring(3);
                        string hexEnd = ((System.Windows.Media.Color)colorEnd).ToString().Substring(3);
                        bpmSideDropHeader.SetStyleAttribute("background", "linear-gradient(#" + hexStart + ",#" + hexEnd + ")");
                        _isBPM_Headerbar_Created = true;
                    }
                }

                // Toggle the toolbar-style for extended view
                CSProfileComponent toolBar = DataAdapter.Instance.DataCache.Profile.Profile_GetComponentFromLayout(Statics.Constants.PROFILE_COMPONENTID_TEMPLATE_TOOLBAR);
                if (toolBar == null)
                {
                    toolBar = DataAdapter.Instance.DataCache.Profile.Profile_GetComponentFromLayout(Statics.Constants.PROFILE_COMPONENTID_TOOLBAR);
                }
                if (toolBar != null)
                {
                    // Setting header-bar - height
                    IVisualObserver obsToolBar = DataAdapter.Instance.FindVisualObserverForProfileComponent(toolBar);
                    double toolbarHeight = obsToolBar.GetPageBindingAdapter().PageProcessing.Page.layoutRoot.ActualHeight;
                    bpmSideDropHeader.SetStyleAttribute("height", toolbarHeight + "px");

                    // Toggle appimage-view
                    ToggleButton btn_appImage = VisualTreeFinder.FindControlDown<ToggleButton>(obsToolBar.GetPageBindingAdapter().PageProcessing.Page.layoutRoot, typeof(ToggleButton), Statics.Constants.IMAGE_NAME_APPIMAGE, -1);
                    if (btn_appImage != null)
                    {
                        btn_appImage.Visibility = originalLayout ? Visibility.Visible : Visibility.Collapsed;
                    }

                    // Change source & visibility
                    bpmSideDropHeaderImage.SetStyleAttribute("display", originalLayout ? "none" : "inline");

                    // Start timer for checking the sideframe-state (if still visible or not)
                    if (!originalLayout)
                    {                        
                        DispatcherTimer timer = new DispatcherTimer();
                        timer.Interval = new TimeSpan(0, 0, 0, 0, 100); // 100 ticks delay
                        timer.Tick += delegate (object s, EventArgs ea)
                        {
                            bool isVisible = false;

                            // Get the display-attribute
                            string displayAttr = HtmlPage.Document.GetElementById("bpmSideDrop").GetStyleAttribute("display");
                            if(displayAttr == null || !displayAttr.ToUpper().Equals("NONE"))
                            {
                                isVisible = true;
                            }

                            if (!isVisible)
                            {
                                // Show the origin-icon
                                btn_appImage = VisualTreeFinder.FindControlDown<ToggleButton>(obsToolBar.GetPageBindingAdapter().PageProcessing.Page.layoutRoot, typeof(ToggleButton), Statics.Constants.IMAGE_NAME_APPIMAGE, -1);
                                btn_appImage.Visibility = Visibility.Visible;

                                // Stop the timer
                                timer.Stop();
                                timer.Tick -= delegate (object s1, EventArgs ea1) { };
                                timer = null;
                            }
                        };
                        timer.Start();
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
        }

        /// <summary>
        /// LF 16.02.2016 - Changed overall for real usage
        /// </summary>
        /// <param name="form"></param>
        /// <param name="div"></param>
        public void BPM_HideSideFrame(HtmlElement form, HtmlElement div)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (form == null) form = HtmlPage.Document.GetElementById("form1");
                if (div == null) div = HtmlPage.Document.GetElementById("bpmSideDrop");
                HtmlElement bpmiFrame = HtmlPage.Document.GetElementById("bpmSideDropFrame");

                if (form != null && div != null)
                {
                    string visible = div.GetStyleAttribute("display");
                    bool musthide = (visible != null && visible.Equals("none"));

                    if (!musthide)
                    {
                        bpmiFrame.SetStyleAttribute("display", "none");
                        div.SetStyleAttribute("display", "none");
                        form.RemoveStyleAttribute("width");
                        form.SetStyleAttribute("width", "100%");
                        form.RemoveStyleAttribute("height");
                        form.SetStyleAttribute("height", "100%");
                    }
                    bpmiFrame.SetProperty("src", "about:blank");
                    bpmiFrame.SetAttribute("currentWF", "");
                    BPM_ToggleToolbarLayout(true);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        #endregion BPM_ToggleSideFrage

        // ==================================================================

        #region Toggle_ShowFullSize (cmd_Toggle_ShowFullSize)

        // TS 08.03.13 speziell für die togglebuttons ohne can_ property damit die buttons nicht disabled werden
        // stattdessen wird die ischecked property an das can_ gebunden über eine zu erstellende Property im BindingAdapter
        //
        // !!! WICHTIG: IM BindingAdapter (und in den Interfaces) muss die CAN_XYZ Property eingetragen werden !!!
        // !!! sonst funktioniert das nicht richtig
        //
        public DelegateCommand cmd_Toggle_ShowFullSize { get { return new DelegateCommand(Toggle_ShowFullSize); } }

        public void Toggle_ShowFullSize()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                //bool processed = false;
                CSProfileComponent profilenode_edesk = DataAdapter.Instance.DataCache.Profile.Profile_GetComponentFromLayout(CSEnumProfileComponentType.EDESKTOP, Constants.PROFILE_COMPONENTID_EDESKTOP);
                CSProfileComponent profilenode_adr = DataAdapter.Instance.DataCache.Profile.Profile_GetComponentFromLayout(CSEnumProfileComponentType.ADRESSEN, Constants.PROFILE_COMPONENTID_ADRESSEN);
                if (((MainPage)App.Current.RootVisual).IsSplittedView)
                {
                    if (!((MainPage)App.Current.RootVisual).IsViewDefault)
                    {
                        if (
                                (((MainPage)App.Current.RootVisual).IsViewAdressen && ViewManager.IsComponentDocked(profilenode_adr))
                            || (((MainPage)App.Current.RootVisual).IsViewDesktop && ViewManager.IsComponentDocked(profilenode_edesk))
                            )
                        {
                            if (!((MainPage)App.Current.RootVisual).IsViewFullSize)
                            {
                                ((MainPage)App.Current.RootVisual).ShowFullSize(((MainPage)App.Current.RootVisual).IsViewAdressen);
                            }
                            else
                                ((MainPage)App.Current.RootVisual).ShowNormalSize();
                        }
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public bool CanToggle_ShowFullSize { get { return _CanToggle_ShowFullSize(); } }

        private bool _CanToggle_ShowFullSize()
        {
            bool ret = DataAdapter.Instance.DataCache.ApplicationFullyInit;
            if (ret)
            {
                //bool processed = false;
                CSProfileComponent profilenode_edesk = DataAdapter.Instance.DataCache.Profile.Profile_GetComponentFromLayout(CSEnumProfileComponentType.EDESKTOP, Constants.PROFILE_COMPONENTID_EDESKTOP);
                CSProfileComponent profilenode_adr = DataAdapter.Instance.DataCache.Profile.Profile_GetComponentFromLayout(CSEnumProfileComponentType.ADRESSEN, Constants.PROFILE_COMPONENTID_ADRESSEN);
                if (((MainPage)App.Current.RootVisual).IsSplittedView)
                {
                    if ((((MainPage)App.Current.RootVisual).IsViewAdressen)
                        || (((MainPage)App.Current.RootVisual).IsViewDesktop))
                    {
                        ret = !((MainPage)App.Current.RootVisual).IsViewFullSize;
                    }
                }
            }
            return ret;
        }

        public bool IsEnabledToggle_ShowFullSize
        {
            get
            {
                bool ret = DataAdapter.Instance.DataCache.ApplicationFullyInit;
                if (ret) ret = ((MainPage)App.Current.RootVisual).IsSplittedView && !((MainPage)App.Current.RootVisual).IsViewDefault;
                return ret;
            }
        }

        #endregion Toggle_ShowFullSize (cmd_Toggle_ShowFullSize)

        // ==================================================================

        #region Toggle_AlwaysLoadAllInTree (cmd_Toggle_AlwaysLoadAllInTree)
        public DelegateCommand cmd_Toggle_AlwaysLoadAllInTree { get { return new DelegateCommand(Toggle_AlwaysLoadAllInTree); } }

        public void Toggle_AlwaysLoadAllInTree()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                bool newVal = !DataAdapter.Instance.DataCache.Profile.Option_GetValueBoolean(CSEnumOptions.treealwaysloadall);
                DataAdapter.Instance.DataCache.Profile.Option_SetValueBoolean(CSEnumOptions.treealwaysloadall, newVal);
                DataAdapter.Instance.InformObservers();

                // If the new state is true (always load all children), then load all children
                if (newVal)
                {
                    IDocumentOrFolder selObj = DataAdapter.Instance.DataCache.Objects(this.Workspace).Folder_Selected;
                    if (selObj != null && selObj.objectId.Length > 0 && !selObj.objectId.Equals("0"))
                    {
                        DataAdapter.Instance.DataCache.Objects(Workspace).ClearCache(false);
                        QuerySingleObjectById(selObj.RepositoryId, selObj.objectId, false, true, CSEnumProfileWorkspace.workspace_default, null);
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public bool CanToggle_AlwaysLoadAllInTree { get { return _CanToggle_AlwaysLoadAllInTree(); } }

        private bool _CanToggle_AlwaysLoadAllInTree()
        {
            bool ret = DataAdapter.Instance.DataCache.ApplicationFullyInit;
            if (ret) ret = DataAdapter.Instance.DataCache.Profile.Option_GetValueBoolean(CSEnumOptions.treealwaysloadall);
            return ret;
        }

        public bool IsEnabledToggle_AlwaysLoadAllInTree
        {
            get
            {
                bool ret = DataAdapter.Instance.DataCache.ApplicationFullyInit;                
                return ret;
            }
        }

        #endregion Toggle_AlwaysLoadAllInTree (cmd_Toggle_AlwaysLoadAllInTree)

        // ==================================================================

        #region TreeViewExpandAll (cmd_TreeViewExpandAll)
        public DelegateCommand cmd_TreeViewExpandAll { get { return new DelegateCommand(TreeViewExpandAll); } }

        public void TreeViewExpandAll()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanTreeViewExpandAll)
                {
                    _treeallloadtriggered = true;
                    DataAdapter.Instance.InformObservers();
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public bool CanTreeViewExpandAll { get { return _CanTreeViewExpandAll(); } }

        private bool _CanTreeViewExpandAll()
        {
            bool ret = DataAdapter.Instance.DataCache.ApplicationFullyInit;
            return ret;
        }

        public bool IsTreeViewExpandAll_Triggered
        {
            get
            {
                return _treeallloadtriggered;
            }
            set
            {
                _treeallloadtriggered = value;
            }
        }

        #endregion Toggle_TreeViewExpandAll (cmd_Toggle_TreeViewExpandAll)

        // ==================================================================

        #region Toggle_Query (cmd_Toggle_Query)

        // TS 08.03.13 speziell für die togglebuttons ohne can_ property damit die buttons nicht disabled werden
        // stattdessen wird die ischecked property an das can_ gebunden über eine zu erstellende Property im BindingAdapter
        //
        // !!! WICHTIG: IM BindingAdapter (und in den Interfaces) muss die CAN_XYZ Property eingetragen werden !!!
        // !!! sonst funktioniert das nicht richtig
        //
        public DelegateCommand cmd_Toggle_Query { get { return new DelegateCommand(Toggle_Query); } }

        public void Toggle_Query()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (CanToggle_Query)
                {
                    Query();
                }
                else
                {
                    ClearCache();
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public bool CanToggle_Query { get { return _CanToggle_Query(); } }

        private bool _CanToggle_Query()
        {
            return CanQuery;
        }

        public bool IsEnabledToggle_Query { get { return DataAdapter.Instance.DataCache.ApplicationFullyInit; } }

        #endregion Toggle_Query (cmd_Toggle_Query)

        // ==================================================================
      
        #region Toggle_LockLayout (cmd_Toggle_LockLayout)

        // TS 08.03.13 speziell für die togglebuttons ohne can_ property damit die buttons nicht disabled werden
        // stattdessen wird die ischecked property an das can_ gebunden über eine zu erstellende Property im BindingAdapter
        //
        // !!! WICHTIG: IM BindingAdapter (und in den Interfaces) muss die CAN_XYZ Property eingetragen werden !!!
        // !!! sonst funktioniert das nicht richtig
        //
        public DelegateCommand cmd_Toggle_LockLayout { get { return new DelegateCommand(Toggle_LockLayout); } }

        public void Toggle_LockLayout()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                CSOption layoutlocked = DataAdapter.Instance.DataCache.Profile.Option_GetOption(CSEnumOptions.layoutlocked);
                bool islayoutlocked = DataAdapter.Instance.DataCache.Profile.Option_GetValueBoolean(CSEnumOptions.layoutlocked);
                layoutlocked.value = (!islayoutlocked).ToString();
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public bool CanToggle_LockLayout { get { return _CanToggle_LockLayout(); } }

        private bool _CanToggle_LockLayout()
        {
            bool ret = DataAdapter.Instance.DataCache.Profile.Option_GetValueBoolean(CSEnumOptions.layoutlocked);
            return !ret;
        }

        public bool IsEnabledToggle_LockLayout { get { return DataAdapter.Instance.DataCache.ApplicationFullyInit; } }

        #endregion Toggle_LockLayout (cmd_Toggle_LockLayout)

        // ==================================================================

        #region Toggle_Options (cmd_Toggle_Options)

        // TS 08.03.13 speziell für die togglebuttons ohne can_ property damit die buttons nicht disabled werden
        // stattdessen wird die ischecked property an das can_ gebunden über eine zu erstellende Property im BindingAdapter
        //
        // !!! WICHTIG: IM BindingAdapter (und in den Interfaces) muss die CAN_XYZ Property eingetragen werden !!!
        // !!! sonst funktioniert das nicht richtig
        //
        public DelegateCommand cmd_Toggle_Options { get { return new DelegateCommand(Toggle_Options); } }

        public void Toggle_Options()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                ShowOptions(null);
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public bool CanToggle_Options { get { return _CanToggle_Options(); } }

        private bool _CanToggle_Options()
        {
            return CanShowOptions;
        }

        public bool IsEnabledToggle_Options { get { return DataAdapter.Instance.DataCache.ApplicationFullyInit; } }

        #endregion Toggle_Options (cmd_Toggle_Options)

        // ==================================================================

        #region Toggle_ShowAnnos (cmd_Toggle_ShowAnnos)

        // TS 08.03.13 speziell für die togglebuttons ohne can_ property damit die buttons nicht disabled werden
        // stattdessen wird die ischecked property an das can_ gebunden über eine zu erstellende Property im BindingAdapter
        //
        // !!! WICHTIG: IM BindingAdapter (und in den Interfaces) muss die CAN_XYZ Property eingetragen werden !!!
        // !!! sonst funktioniert das nicht richtig
        //
        public DelegateCommand cmd_Toggle_ShowAnnos { get { return new DelegateCommand(Toggle_ShowAnnos); } }

        public void Toggle_ShowAnnos()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                CSOption showannos = DataAdapter.Instance.DataCache.Profile.Option_GetOption(CSEnumOptions.viewershowannos);
                bool isannosvisible = DataAdapter.Instance.DataCache.Profile.Option_GetValueBoolean(CSEnumOptions.viewershowannos);
                showannos.value = (!isannosvisible).ToString();
                DataAdapter.Instance.InformObservers();
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public bool CanToggle_ShowAnnos { get { return _CanToggle_ShowAnnos(); } }

        private bool _CanToggle_ShowAnnos()
        {
            bool ret = DataAdapter.Instance.DataCache.Profile.Option_GetValueBoolean(CSEnumOptions.viewershowannos);
            return !ret;
        }

        public bool IsEnabledToggle_ShowAnnos
        {
            get
            {
                bool ret = DataAdapter.Instance.DataCache.ApplicationFullyInit;
                if (ret) ret = DataAdapter.Instance.DataCache.Objects(this.Workspace).Document_Selected.objectId.Length > 0;
                return ret;
            }
        }

        #endregion Toggle_ShowAnnos (cmd_Toggle_ShowAnnos)

        // ==================================================================

        #region Toggle_Logout (cmd_Toggle_Logout)

        // TS 08.03.13 speziell für die togglebuttons ohne can_ property damit die buttons nicht disabled werden
        // stattdessen wird die ischecked property an das can_ gebunden über eine zu erstellende Property im BindingAdapter
        //
        // !!! WICHTIG: IM BindingAdapter (und in den Interfaces) muss die CAN_XYZ Property eingetragen werden !!!
        // !!! sonst funktioniert das nicht richtig
        //
        public DelegateCommand cmd_Toggle_Logout { get { return new DelegateCommand(Toggle_Logout); } }

        public void Toggle_Logout()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (!DataAdapter.Instance.DataCache.Objects(this.Workspace).ObjectList_InWork.View.IsEmpty)
                {
                    // meldung ausgeben
                    DisplayWarnMessage(LocalizationMapper.Instance["msg_warn_unsaved"]);
                    // das hier muß sein damit der button nicht einfach umschaltet (wird hiermit automatisch zurückgesetzt bzw. einfach refreshed)
                    DataAdapter.Instance.InformObservers(this.Workspace);
                }
                else
                {

                    if (isMasterApplication)
                    {
                        Logout_WriteProfile();
                    }else
                    {
                        // Logout is only allowed on the master-applikation
                        DisplayWarnMessage(LocalizationMapper.Instance["msg_warn_logoutonlyonmaster"]);
                        DataAdapter.Instance.InformObservers(this.Workspace);
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public bool CanToggle_Logout { get { return _CanToggle_Logout(); } }

        private bool _CanToggle_Logout()
        {
            //bool ret = DataAdapter.Instance.DataCache.Profile.Option_GetValueBoolean(CSEnumOptions.annosvisible);
            //return !ret;
            return DataAdapter.Instance.DataCache.ApplicationFullyInit;
        }

        public bool IsEnabledToggle_Logout { get { return DataAdapter.Instance.DataCache.ApplicationFullyInit; } }

        private void Logout_WriteProfile()
        {
            // TS 11.11.15
            //this.Profile_UpdateLayoutComponents("", new CallbackAction(Logout_LCInform));
            this.Profile_UpdateLayoutComponents("", new CallbackAction(Logout_LCInformed));
        }

        // TS 11.11.15 wird nicht mehr so mitgeteilts sondern der server sendet sowieso ein logoutcommitted an die queue !!
        //private void Logout_LCInform()
        //{
        //    // TS 22.07.14 auch lc informieren
        //    string lcexecutefile = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.LC_Command];
        //    string lcparamfilename = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.LC_Param];
        //    CSLCRequest lcrequest = CreateLCRequest();
        //    bool pushtolc = LC_IsPushEnabled;

        //    DataAdapter.Instance.DataProvider.LC_CreateCommandFile(CSEnumCommands.cmd_Toggle_Logout, lcparamfilename, lcrequest, pushtolc, DataAdapter.Instance.DataCache.Rights.UserPrincipal, new CallbackAction(Logout_LCInformed));
        //    // LC nur rufen nur wenn nicht bereits beim serveraufruf gepushed wurde
        //    if (!pushtolc && lcexecutefile != null && lcexecutefile.Length > 0 && lcparamfilename != null && lcparamfilename.Length > 0)
        //    {
        //        LC_ExecuteCall(lcexecutefile, lcparamfilename);
        //    }
        //}
        private void Logout_LCInformed()
        {
            // TS 23.02.15 erst die traces noch wegschreiben
            Log.Log.ForceFlush(new CallbackAction(Logout_FlushTracesDone));

            // TS 23.02.15
            //DataAdapter.Instance.DataProvider.Logout(DataAdapter.Instance.DataCache.Rights.UserPrincipal, new CallbackAction(LogoutDone));
        }

        private void Logout_FlushTracesDone()
        {
            // dann erst die session killen
            DataAdapter.Instance.DataProvider.Logout(DataAdapter.Instance.DataCache.Rights.UserPrincipal, new CallbackAction(LogoutDone));
        }

        private void LogoutDone()
        {
            if (DataAdapter.Instance.DataCache.ResponseStatus.success)
            {
                ClientCommunication.Communicator.Instance.SendMessageToAll(EnumHelper.GetValueFromCSEnum(Constants.EnumCommunicatorMessageTypes.LOGOUT_MASTERPAGE));
                // TS 02.09.15 reihenfolge getauscht da die sessionguid beim viewmanager.unloadall abgefragt wird und wenn sie leer ist wird nicht ordentlich entladen
                //DataAdapter.Instance.DataCache.Rights.UserPrincipal.sessionguid = "";
                BPM_HideSideFrame(null, null);
                AppManager.UnloadAll(true);

                // TS 02.09.15 reihenfolge getauscht da die sessionguid beim viewmanager.unloadall abgefragt wird und wenn sie leer ist wird nicht ordentlich entladen
                DataAdapter.Instance.DataCache.Rights.UserPrincipal.sessionguid = "";

                // Close HTML-Desktop
                Init.ViewManager.CloseHTMLDesktop();
            }
        }

        #endregion Toggle_Logout (cmd_Toggle_Logout)

        // ==================================================================

        #region Toggle_Show_Help (cmd_Toggle_Show_Help)

        // TS 08.03.13 speziell für die togglebuttons ohne can_ property damit die buttons nicht disabled werden
        // stattdessen wird die ischecked property an das can_ gebunden über eine zu erstellende Property im BindingAdapter
        //
        // !!! WICHTIG: IM BindingAdapter (und in den Interfaces) muss die CAN_XYZ Property eingetragen werden !!!
        // !!! sonst funktioniert das nicht richtig
        //
        public DelegateCommand cmd_Toggle_Help { get { return new DelegateCommand(Toggle_Help); } }

        public void Toggle_Help()
        {
            //if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                ShowHelp();
            }
            catch (Exception e) { Log.Log.Error(e); }
            //if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public bool CanToggle_Help { get { return _CanToggle_Help(); } }

        private bool _CanToggle_Help()
        {
            return DataAdapter.Instance.DataCache.ApplicationFullyInit;
        }

        public bool IsEnabledToggle_Help { get { return DataAdapter.Instance.DataCache.ApplicationFullyInit; } }

        #endregion Toggle_Show_Help (cmd_Toggle_Show_Help)

        // ==================================================================

        #region Toggle_ShowFavs (cmd_Toggle_ShowFavs)

        // TS 08.03.13 speziell für die togglebuttons ohne can_ property damit die buttons nicht disabled werden
        // stattdessen wird die ischecked property an das can_ gebunden über eine zu erstellende Property im BindingAdapter
        //
        // !!! WICHTIG: IM BindingAdapter (und in den Interfaces) muss die CAN_XYZ Property eingetragen werden !!!
        // !!! sonst funktioniert das nicht richtig
        //
        public DelegateCommand cmd_Toggle_ShowFavs { get { return new DelegateCommand(Toggle_ShowFavs); } }

        public void Toggle_ShowFavs_Hotkey() { Toggle_ShowFavs(); }

        public void Toggle_ShowFavs()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                Favs_Show();
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public bool CanToggle_ShowFavs { get { return _CanToggle_ShowFavs(); } }

        private bool _CanToggle_ShowFavs()
        {
            return DataAdapter.Instance.DataCache.ApplicationFullyInit;
        }

        public bool IsEnabledToggle_ShowFavs { get { return DataAdapter.Instance.DataCache.ApplicationFullyInit; } }

        #endregion Toggle_ShowFavs (cmd_Toggle_ShowFavs)

        // ==================================================================

        #region diverse publics
        public void DisplayMessage(string message)
        {
            try
            {
                CSResponseStatus status = new CSResponseStatus();
                status.success = true;
                status.returncode = Constants.RESPONSESTATUS_CODE_WARNING;
                status.localizedmessage = message;
                ((IDataObserver)DataAdapter.Instance).ApplyData(status);
            }
            catch (Exception) { }
        }


        public void DisplayWarnMessage(string message)
        {
            try
            {
                bool displaywarnings = DataAdapter.Instance.DataCache.Profile.Option_GetValueBoolean(CSEnumOptions.displaywarnings);
                if (displaywarnings) ViewManager.ShowTimedWarnMessage(message);
                // und zusätzlich im statustext ausgeben
                CSResponseStatus status = new CSResponseStatus();
                // TS 10.04.14
                //status.success = false;
                status.success = true;
                status.returncode = Constants.RESPONSESTATUS_CODE_WARNING;
                //status.localizedmessage = Localization.localstring.msg_autosave_document_notdone;
                status.localizedmessage = message;
                ((IDataObserver)DataAdapter.Instance).ApplyData(status);
            }
            catch (Exception) { }
        }

        /// <summary>
        /// Navigate to the object with the given id, if available in datacache-default
        /// TODO: Actually this method only takes care of folder-links!
        /// </summary>
        /// <param name="objID"></param>
        public void displayObjectFromRCLink(object cmdParams, object objWorkspace)
        {
            try
            {
                Dictionary<string, string> paramDict = (Dictionary<string, string>)cmdParams;
                CSEnumProfileWorkspace enWorkspace = (CSEnumProfileWorkspace)objWorkspace;
                string objectID = "";

                if (DataAdapter.Instance.DataCache.ResponseStatus.success)
                {
                    // Read Folder-Link
                    if (paramDict.TryGetValue(Statics.ExternalCommandHelper.ExternalCommands.FOLDER.ToString(), out objectID))
                    {
                        IDocumentOrFolder selObj = DataAdapter.Instance.DataCache.Objects(enWorkspace).GetObjectById(objectID);
                        if (selObj != null && selObj.objectId.Length > 0)
                        {
                            DataAdapter.Instance.DataCache.Objects(enWorkspace).SetSelectedObject(selObj);
                            RefreshDZCollectionsAfterProcess(selObj, "");
                            DataAdapter.Instance.NavigateUI(enWorkspace, enWorkspace);
                            ((MainPage)App.Current.RootVisual).ShowDefault();
                        }
                    }
                }
            }
            catch (Exception ex) { if (ex.Message.Length > 0) { Log.Log.Debug(ex.Message); } }
        }

        /// <summary>
        /// wird gerufen nach copy, paste, move hat noch eigenen handler
        /// </summary>
        /// <param name="wasdocument"></param>
        public void ObjectTransfer_Done(bool wasdocument)
        {
            if (DataAdapter.Instance.DataCache.ResponseStatus.success)
            {
                DataAdapter.Instance.DataCache.Info.ListView_ForcePageSelection = true;
                if (Workspace == CSEnumProfileWorkspace.workspace_clipboard)
                {
                    DataAdapter.Instance.InformObservers(Workspace);
                }
                else
                {
                    if (wasdocument && DataAdapter.Instance.DataCache.Objects(Workspace).ObjectId_LastAddedDocument.Length > 0)
                    {
                        IDocumentOrFolder obj = DataAdapter.Instance.DataCache.Objects(Workspace).GetObjectById(DataAdapter.Instance.DataCache.Objects(Workspace).ObjectId_LastAddedDocument);
                        SetSelectedObject(obj.objectId);
                    }
                    else if (DataAdapter.Instance.DataCache.Objects(Workspace).ObjectIdList_LastAddedFolders.Count > 0)
                    {
                        string objectid = DataAdapter.Instance.DataCache.Objects(Workspace).ObjectIdList_LastAddedFolders[DataAdapter.Instance.DataCache.Objects(Workspace).ObjectIdList_LastAddedFolders.Count - 1];
                        //if (!DataAdapter.Instance.DataCache.Objects(Workspace).Folder_Selected.objectId.Equals(objectid))
                            SetSelectedObject(objectid);
                    }
                }
            }
        }
        public void PushLCCheckToListeners()
        {
            // Check the LC-Availability and return the push to all listeners
            bool connectorAvailable = DataAdapter.Instance.Processing(CSEnumProfileWorkspace.workspace_default).LC_IsAvailable;
            connectorAvailable = connectorAvailable && DataAdapter.Instance.Processing(CSEnumProfileWorkspace.workspace_default).LC_IsModulAvailable(CSEnumInformationId.LC_Modul_Word, CSEnumInformationId.LC_Modul_WordParam);

            List<string> commandsID = new List<string>();
            commandsID.Add(connectorAvailable.ToString());

            CSRCPushPullUIRouting routinginfo = new CSRCPushPullUIRouting();
            CSEnumRCPushUIListeners?[] receiver = new CSEnumRCPushUIListeners?[2];
            receiver[0] = CSEnumRCPushUIListeners.HTML_CLIENT_DESKTOP;
            receiver[1] = CSEnumRCPushUIListeners.HTML_CLIENT_ADRESS;
            routinginfo.sendto_componentorvirtualids = receiver;
            DataAdapter.Instance.DataProvider.RCPush_UI(routinginfo, CSEnumRCPushCommands.checkconnector, commandsID, new List<cmisObjectType>(), DataAdapter.Instance.DataCache.Rights.UserPrincipal, null);
        }

        public void RefreshDZCollectionsAfterProcess(object objectorlist, string mustnotsetid)
        {
            if (objectorlist != null && objectorlist.GetType().Name.StartsWith("List"))
            {
                List<IDocumentOrFolder> objectlist = (List<IDocumentOrFolder>)objectorlist;
                RefreshDZCollectionsAfterProcess(objectlist, mustnotsetid);
            }
            else if (objectorlist != null && objectorlist.GetType().Name.StartsWith("DocumentOrFolder"))
            {
                IDocumentOrFolder docorfolder = (IDocumentOrFolder)objectorlist;
                RefreshDZCollectionsAfterProcess(docorfolder, mustnotsetid);
            }
        }

        public void RefreshDZCollectionsAfterProcess(IDocumentOrFolder docorfolder, string mustnotsetid)
        {
            List<IDocumentOrFolder> objectlist = new List<IDocumentOrFolder>();
            objectlist.Add(docorfolder);
            RefreshDZCollectionsAfterProcess(objectlist, mustnotsetid);
        }

        /// <summary>
        /// ermittelt distinct alle workspaces von in der liste enthaltenen dokumenten
        /// und aktualisiert dort das aktuell gewählte doklog (wenn denn eines gewählt ist oder eine datei darunter)
        /// </summary>
        /// <param name="objectlist"></param>
        public void RefreshDZCollectionsAfterProcess(List<IDocumentOrFolder> objectlist, string mustnotsetid)
        {
            List<CSEnumProfileWorkspace> workspaces = new List<CSEnumProfileWorkspace>();
            bool relationsChanged = false;

            foreach (IDocumentOrFolder obj in objectlist)
            {
                // TS 25.02.14 auch bei logischen dokumente muss die collection aktualisiert werden
                obj.NotifyPropertyChanged(CSEnumCmisProperties.workspace.ToString());
                if (obj.isDocument || obj.structLevel == Constants.STRUCTLEVEL_09_DOKLOG)
                {
                    if (!workspaces.Contains(obj.Workspace))
                    {
                        workspaces.Add(obj.Workspace);
                    }
                }
                relationsChanged = obj.hasRelationships ? true : relationsChanged;
            }

            List<string> selecteddones = new List<string>();
            if (workspaces.Count > 0)
            {
                foreach (CSEnumProfileWorkspace workspace in workspaces)
                {
                    IDocumentOrFolder currentselected = DataAdapter.Instance.DataCache.Objects(workspace).Folder_Selected;

                    // In case of workspace is "aufgabe" and the selected folder is the root object, don't do anything except a refresh
                    bool isTaskRoot = false;
                    if (workspace == CSEnumProfileWorkspace.workspace_aufgabe && DataAdapter.Instance.DataCache.Objects(workspace).Root.objectId.Equals(currentselected.objectId))
                    {
                        isTaskRoot = true;
                        DataAdapter.Instance.InformObservers();
                    }
                    if (currentselected.hasChildDocuments && !isTaskRoot)
                    {
                        // TS 02.07.15 die mustnotsetid weglassen
                        // DataAdapter.Instance.Processing(workspace).GetObjectsUpDown(currentselected, false, null);
                        if (mustnotsetid == null || mustnotsetid.Length == 0 || !currentselected.objectId.Equals(mustnotsetid))
                        {
                            DataAdapter.Instance.Processing(workspace).GetObjectsUpDown(currentselected, false, null);
                        }
                        if (!selecteddones.Contains(currentselected.objectId))
                            selecteddones.Add(currentselected.objectId);
                    }
                }
            }
            // und nochmal durchgehen für die "echten" MoveParents, damit nach einem cancel z.b. dort nicht die bilder stehen bleiben
            foreach (IDocumentOrFolder obj in objectlist)
            {
                if (obj.isDocument && obj.TempOrRealParentId != null && obj.TempOrRealParentId.Length > 0 && !selecteddones.Contains(obj.TempOrRealParentId))
                {
                    // workspace suchen
                    foreach (CSEnumProfileWorkspace ws in DataAdapter.Instance.DataCache.WorkspacesUsed)
                    {
                        if (DataAdapter.Instance.DataCache.Objects(ws).ExistsObject(obj.TempOrRealParentId))
                        {
                            DataAdapter.Instance.Processing(ws).GetObjectsUpDown(DataAdapter.Instance.DataCache.Objects(ws).GetObjectById(obj.TempOrRealParentId), false, null);
                            break;
                        }
                    }
                }

                // Update children after move
                if (obj[CSEnumCmisProperties.tmp_saveWasMoved] != null && obj[CSEnumCmisProperties.tmp_saveWasMoved] == "true")
                {
                    DataAdapter.Instance.Processing(obj.Workspace).GetObjectsUpDown(obj, true, null);
                }
            }

            // If some relations may have changed, refresh ui-elements for link-symbols
            if (relationsChanged)
            {
                DispatcherTimer timer = new DispatcherTimer();
                timer.Interval = new TimeSpan(0, 0, 0, 0, 20);
                timer.Tick += delegate (object s, EventArgs ea)
                {
                    timer.Stop();

                    UpdateRelationshipsForObjects(objectlist);
                    timer.Tick -= delegate (object s1, EventArgs ea1) { };
                    timer = null;
                };
                timer.Start();                
            }
        }

        /// <summary>
        /// Update all related objects from the given object-list
        /// </summary>
        /// <param name="origObjects"></param>
        private void UpdateRelationshipsForObjects(List<IDocumentOrFolder> origObjects)
        {
            List<string> relationList = new List<string>();
            List<IDocumentOrFolder> relObjList = new List<IDocumentOrFolder>();

            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                // Clean orphan relationships                
                foreach (IDocumentOrFolder obj in origObjects)
                {
                    
                    foreach(KeyValuePair<DocOrFolderRelationship, cmisObjectType> relpair in obj.deletableRelationships)
                    {
                        List<cmisObjectType> delObj = new List<cmisObjectType>();
                        
                        bool isTargetInCache = DataAdapter.Instance.DataCache.Object_FindObjectWithID(relpair.Key.target_objectid).Count > 0;
                        bool isSourceInCache = DataAdapter.Instance.DataCache.Object_FindObjectWithID(relpair.Key.source_objectid).Count > 0; 

                        delObj.Add(relpair.Value);
                        DataAdapter.Instance.DataProvider.RemoveRelationship(delObj, this.Workspace, isTargetInCache, isSourceInCache, DataAdapter.Instance.DataCache.Rights.UserPrincipal, new CallbackAction(RemoveRelationship_Done, delObj));
                    }                    
                }

                // Get objects with relations
                foreach (IDocumentOrFolder origObj in origObjects)
                {
                    if (origObj.hasRelationships == true)
                    {
                        relationList.AddRange(origObj.getRelationshipTargets);
                    }
                }

                // Search for related objects in all workspaces
                foreach (string rel in relationList)
                {
                    relObjList.AddRange(DataAdapter.Instance.DataCache.Object_FindObjectWithID(rel));
                }

                // And update the gathered data
                foreach (IDocumentOrFolder relObj in relObjList)
                {
                    if (!relObj.Workspace.ToString().ToUpper().EndsWith("_RELATED") && !relObj.WorkspaceAddList.Contains(CSEnumProfileWorkspace.workspace_links) && !relObj.WorkspaceAddList.Contains(CSEnumProfileWorkspace.workspace_favoriten) && !relObj.WorkspaceAddList.Contains(CSEnumProfileWorkspace.workspace_clipboard))
                    {
                        DataAdapter.Instance.DataProvider.GetObject(relObj.RepositoryId, relObj.objectId, relObj.Workspace, DataAdapter.Instance.DataCache.Rights.UserPrincipal, new CallbackAction(UpdateRelationshipsForObjects_Done, relObj));
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }


        private void RemoveRelationship_Done(object relationObjects)
        {
            // Update Treeview-Elements
            List<cmisObjectType> relList = (List<cmisObjectType>)relationObjects;
            if(relList.Count > 0)
            {
                foreach(cmisObjectType relObj in relList)
                {
                    // Get the source-object from the datacache
                    IDocumentOrFolder sourceObj = null;
                    string sourceID = DiverseHelpers.GetCMISPropertyValueAsString(relObj, CSEnumCmisProperties.cmis_sourceId, 0);
                    if(sourceID.Length > 0)
                    {
                        sourceObj = DataAdapter.Instance.DataCache.Objects(this.Workspace).GetObjectById(sourceID);
                    }

                    // Get the target-object from the datacache
                    IDocumentOrFolder targetObj = null;
                    string targetID = DiverseHelpers.GetCMISPropertyValueAsString(relObj, CSEnumCmisProperties.cmis_targetId, 0);
                    if (targetID.Length > 0)
                    {
                        targetObj = DataAdapter.Instance.DataCache.Objects(this.Workspace).GetObjectById(targetID);
                    }

                    // Maybe update those objects? => SOURCE
                    if(sourceObj != null && sourceObj.objectId.Length > 0)
                    {
                        sourceObj.NotifyPropertyChanged("");
                    }

                    // Maybe update those objects? => TARGET
                    if (targetObj != null && targetObj.objectId.Length > 0)
                    {
                        targetObj.NotifyPropertyChanged("");
                    }
                }
            }

            DispatcherTimer timer = new DispatcherTimer();
            timer.Interval = new TimeSpan(0, 0, 0, 0, 20); // 100 ticks delay for updating UI
            timer.Tick += delegate (object s, EventArgs ea)
            {
                timer.Stop();
                timer.Tick -= delegate (object s1, EventArgs ea1) { };
                DataAdapter.Instance.InformObservers();
                timer = null;
            };
            timer.Start();
        }

        /// <summary>
        /// Fire NotifyPropertyChanged-Events for the updated objects
        /// </summary>
        /// <param name="updatedObject"></param>
        private void UpdateRelationshipsForObjects_Done(object updatedObject)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                DispatcherTimer timer = new DispatcherTimer();
                timer.Interval = new TimeSpan(0, 0, 0, 0, 20); // 100 ticks delay for updating UI
                timer.Tick += delegate (object s, EventArgs ea)
                {
                    timer.Stop();
                    timer.Tick -= delegate (object s1, EventArgs ea1) { };
                    IDocumentOrFolder obj = (IDocumentOrFolder)updatedObject;
                    if (obj != null && obj.objectId.Length > 0)
                    {
                        obj.NotifyPropertyChanged(CSEnumCmisProperties.cmis_targetId.ToString());
                        DataAdapter.Instance.InformObservers(obj.Workspace);
                    }
                    timer = null;
                };
                timer.Start();

            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        public void InitExtendedQueryValues()
        {
            _queryext_musthavevalues = new ExtendedQueryValues(this.Workspace);
            _queryext_shouldhavevalues = new ExtendedQueryValues(this.Workspace);
            _queryext_mustnothavevalues = new ExtendedQueryValues(this.Workspace);
        }

        public void ExtendedQueryExtractValues(Dictionary<string, string> propertynames, Dictionary<string, string> propertyoperators)
        {
            Dictionary<int, Dictionary<string, string>> alllevelvalues = new Dictionary<int, Dictionary<string, string>>();
            for (int i = 0; i < 3; i++)
            {
                ExtendedQueryValues extvals = QueryShouldHaveValues;

                bool multior = false;
                if (i == 0)
                {
                    // prüfen ob mehrere felder betroffen (verODERt) sind
                    string lastval = "";
                    string currval = "";
                    foreach (ExtendedQueryValue eqv in extvals.Values)
                    {
                        currval = eqv.PropertyName;
                        if (lastval.Length == 0)
                            lastval = currval;
                        else
                        {
                            if (!lastval.Equals(currval))
                            {
                                multior = true;
                                break;
                            }
                        }
                    }
                }
                if (i == 1)
                    extvals = QueryMustHaveValues;
                else if (i == 2)
                    extvals = QueryMustNotHaveValues;

                foreach (ExtendedQueryValue eqv in extvals.Values)
                {
                    string propdisplayname = eqv.PropertyName;
                    string propopdisplayname = eqv.PropertyOperator;
                    string propval = eqv.PropertyValue;
                    string oldval = "";
                    string propqueryname = "";
                    string propopqueryname = "";

                    // propertyname von display auf query mappen
                    foreach (KeyValuePair<string, string> kvp in propertynames)
                    {
                        if (kvp.Value.Equals(propdisplayname))
                        {
                            propqueryname = kvp.Key;
                            break;
                        }
                    }
                    // operator von display auf query mappen propertyoperators
                    foreach (KeyValuePair<string, string> kvp in propertyoperators)
                    {
                        //string propdisplayname = eqv.PropertyName;
                        if (kvp.Value.Equals(propopdisplayname))
                        {
                            propopqueryname = kvp.Key;
                            break;
                        }
                    }

                    cmisTypeContainer cont = DataAdapter.Instance.DataCache.Repository(this.Workspace).GetTypeContainerForPropertyName(propqueryname);
                    if (cont != null)
                    {
                        string typeid = cont.type.id;
                        int findlevel = StructLevelFinder.GetStructLevelFromTypeId(typeid);
                        if (findlevel > 0)
                        {
                            if (!alllevelvalues.ContainsKey(findlevel))
                            {
                                alllevelvalues.Add(findlevel, new Dictionary<string, string>());
                            }
                            Dictionary<string, string> levelvalues = alllevelvalues[findlevel];

                            if (levelvalues.ContainsKey(propqueryname))
                            {
                                oldval = levelvalues[propqueryname];
                                levelvalues.Remove(propqueryname);
                            }
                            if (oldval.Length > 0)
                            {
                                oldval = oldval + " ";
                                if (i == 0) oldval = oldval + Constants.QUERY_OPERATOR_OR;
                            }
                            else
                            {
                                if (i == 0 && multior) oldval = Constants.QUERY_OPERATOR_OR;
                            }
                            // operator davor wenn nicht die ausschlussliste, dann wird automatische ein != davor gemacht
                            if (i < 2 && propopqueryname.Length > 0)
                                oldval = oldval + propopqueryname;
                            else if (i == 2)
                                oldval = oldval + "!=";

                            oldval = oldval + propval;
                            levelvalues.Add(propqueryname, oldval);
                        }
                    }
                }
            }

            foreach (KeyValuePair<int, Dictionary<string, string>> kvp in alllevelvalues)
            {
                if (kvp.Value.Count > 0)
                {
                    foreach (KeyValuePair<string, string> kvpdetail in kvp.Value)
                    {
                        DataAdapter.Instance.DataCache.Objects(this.Workspace).Folder_Selected_Level(kvp.Key)[kvpdetail.Key] = kvpdetail.Value;                        
                        //Page.PageBindingAdapter.DataCacheObjects.Folder_Selected_Level(kvp.Key)[kvpdetail.Key] = kvpdetail.Value;
                        //DataCacheObjects.Folder_Selected_Level(kvp.Key)[kvpdetail.Key] = kvpdetail.Value;
                    }
                }
            }
        }

        public void ExtendedQueryRemoveValues(Dictionary<string, string> propertynames, string propdisplayname)
        {
            Dictionary<int, Dictionary<string, string>> alllevelvalues = new Dictionary<int, Dictionary<string, string>>();
            // propertyname von display auf query mappen
            string propqueryname = "";
            foreach (KeyValuePair<string, string> kvp in propertynames)
            {
                if (kvp.Value.Equals(propdisplayname))
                {
                    propqueryname = kvp.Key;
                    break;
                }
            }
            if (propqueryname != null && propqueryname.Length > 0)
            {
                cmisTypeContainer cont = DataAdapter.Instance.DataCache.Repository(this.Workspace).GetTypeContainerForPropertyName(propqueryname);
                if (cont != null)
                {
                    string typeid = cont.type.id;
                    int findlevel = StructLevelFinder.GetStructLevelFromTypeId(typeid);
                    if (findlevel > 0)
                    {

                        DataAdapter.Instance.DataCache.Objects(this.Workspace).Folder_Selected_Level(findlevel)[propqueryname] = "";
                    }
                }
            }
        }

        public void ExtendedQueryClearAllValues()
        {
            _queryext_musthavevalues.Values.Clear();
            _queryext_shouldhavevalues.Values.Clear();
            _queryext_mustnothavevalues.Values.Clear();
        }

        public void ExtendedQueryStartQuery()
        {
            // TS 17.07.17
            //_queryext_musthavevalues.RefreshSearchConditions();
            //_queryext_shouldhavevalues.RefreshSearchConditions();
            //_queryext_mustnothavevalues.RefreshSearchConditions();
            //this.Query();
            this.ClearCache(false, new CallbackAction(ExtendedQueryStartQuery_Callback));
        }

        private void ExtendedQueryStartQuery_Callback()
        {
            _queryext_musthavevalues.RefreshSearchConditions();
            _queryext_shouldhavevalues.RefreshSearchConditions();
            _queryext_mustnothavevalues.RefreshSearchConditions();
            this.Query();
        }

        #region Delete-Request

        public void TriggerDeleteRequest(IDocumentOrFolder trigger_object)
        {
            List<IDocumentOrFolder> checkoutList = new List<IDocumentOrFolder>();
            checkoutList.Add(trigger_object);

            // Checkout the object, and notify the desktop in the callback
            SetCheckoutState(checkoutList, true, Constants.SPEC_PROPERTY_STATUS_DELETEREQUEST, new CallbackAction(TriggerDeleteRequest_PushToDesktop, trigger_object));
        }

        private void TriggerDeleteRequest_PushToDesktop(object trigger_object)
        {
            List<string> objectidlist = new List<string>();
            List<cmisObjectType> cmisobjectlist = null;
            IDocumentOrFolder triggerObject = (IDocumentOrFolder)trigger_object;

            // Add the objects display-id
            objectidlist.Add("2C_" + triggerObject.RepositoryId + "_" + triggerObject.objectId);            

            CSRCPushPullUIRouting routinginfo = new CSRCPushPullUIRouting();
            CSEnumRCPushUIListeners?[] receiver = new CSEnumRCPushUIListeners?[1];
            receiver[0] = CSEnumRCPushUIListeners.HTML_CLIENT_DESKTOP;
            routinginfo.sendto_componentorvirtualids = receiver;

            

            // If the HTML-Desktop is already open, send a push-command
            if (Init.ViewManager.IsHTMLDesktopOpen())
            {
                DataAdapter.Instance.DataProvider.RCPush_UI(routinginfo, CSEnumRCPushCommands.createdeleterequest, objectidlist, cmisobjectlist, DataAdapter.Instance.DataCache.Rights.UserPrincipal);
            }
            else
            {
                // Open the desktop-tool                        
                Toggle_ShowDesk(CSEnumRCPushCommands.createdeleterequest.ToString() + "=" + objectidlist[0]);
            }
        }
        #endregion


        #endregion diverse publics

        // ******************************************************************
        // INTERNE PRIVATE COMMANDS OHNE CMD_ Command-Bindung
        // ******************************************************************

        #region diverse privates

        private void _InitWorkspaceQueryOrders()
        {
            CSProfileWorkspace ws = DataAdapter.Instance.DataCache.Profile.Profile_GetProfileWorkspaceFromType(this.Workspace);
            if (ws != null && ws.datafilters != null)
            {
                foreach (CSProfileDataFilter df in ws.datafilters)
                {
                    if (df.type == CSEnumProfileDataFilterType.ws_query_orderby)
                    {
                        CSOrderToken orderby = new CSOrderToken();
                        orderby.orderby = CSEnumOrderBy.asc;
                        string propertyname = df.selectedvalues;
                        if (propertyname.Contains(" "))
                        {
                            string ascdesc = propertyname.Substring(propertyname.IndexOf(" ") + 1);
                            propertyname = propertyname.Substring(0, propertyname.IndexOf(" "));
                            if (ascdesc.ToUpper().Equals("DESC"))
                            {
                                orderby.orderby = CSEnumOrderBy.desc;
                                // TS 28.06.13 das muss hier dazu sonst nimmt er immer den default der klasse = asc
                                orderby.orderbySpecified = true;
                            }
                        }
                        orderby.propertyname = propertyname;
                        WorkspaceOrderTokens.Add(orderby);
                    }
                }
            }
        }

        private void _CollectDZImagesRecursive(IDocumentOrFolder docorfolder, List<CServer.cmisObjectType> cmisdocuments)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                for (int i = 0; i < docorfolder.ChildDocuments.Count; i++)
                {
                    IDocumentOrFolder doc = docorfolder.ChildDocuments[i];
                    cmisdocuments.Add(doc.CMISObject);
                    if (doc.ChildDocuments.Count > 0)
                        _CollectDZImagesRecursive(doc, cmisdocuments);
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        private void _CreateFoldersInternal(CSEnumInternalObjectType objecttype, CSEnumProfileWorkspace workspace, int structlevel, bool showdialog)
        {
            _CreateFoldersInternal(objecttype, workspace, structlevel, DataAdapter.Instance.DataCache.Objects(this.Workspace).Objects_Selected, showdialog);
        }

        private void _CreateFoldersInternal(CSEnumInternalObjectType objecttype, CSEnumProfileWorkspace workspace, int structlevel, List<IDocumentOrFolder> selectedobjects, bool showdialog)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                // stapel hat eigene funktion
                if (_CanCreateFoldersInternal(objecttype, workspace, structlevel, selectedobjects))
                {
                    cmisTypeContainer typedef = DataAdapter.Instance.DataCache.Repository(workspace).GetTypeContainerForStructLevel(structlevel);

                    cmisObjectType parentobject = DataAdapter.Instance.DataCache.Objects(workspace).Root.CMISObject;
                    CServer.cmisTypeContainer typedefrelation = null;
                    List<cmisObjectType> relationobjects = new List<cmisObjectType>();
                    bool autosave = false;
                    // TS 10.11.14
                    if (!showdialog)
                        autosave = true;

                    if (selectedobjects != null && selectedobjects.Count > 0)
                    {
                        typedefrelation = DataAdapter.Instance.DataCache.Repository(Workspace).GetTypeContainerForBaseId("cmisrelationship");
                        // TS 28.02.14
                        if (typedefrelation == null) typedefrelation = DataAdapter.Instance.DataCache.Repository_FindTypeContainerForBaseId("cmisrelationship");

                        // zur zeit wird nur das erste objekt verarbeitet, zunächst muss die komplette logik für andere user erweitert werden !
                        // TS 09.10.14 umbau auf mehrere (wieder rausgenommen, da die indizes nicht für alle gesetzt werden aus dem dialog)

                        // TS 09.02.15 nur wenn nicht eine aufgabe oder termin als relation mitgegeben wurde
                        CSEnumProfileWorkspace ws = selectedobjects[0].Workspace;
                        if (!ws.ToString().Equals(CSEnumProfileWorkspace.workspace_aufgabe.ToString()) && !ws.ToString().Equals(CSEnumProfileWorkspace.workspace_termin.ToString()))
                        {
                            relationobjects.Add(selectedobjects[0].CMISObject);
                        }

                        //foreach (IDocumentOrFolder obj in selectedobjects)
                        //{
                        //    relationobjects.Add(obj.CMISObject);
                        //}
                    }

                    //TODO: .SelectedPage fehlt noch

                    if (DataAdapter.Instance.DataCache.Objects(workspace).Object_Selected.objectId.Length == 0)
                    {
                        DataAdapter.Instance.DataCache.Objects(workspace).SetSelectedObject(DataAdapter.Instance.DataCache.Objects(workspace).Root);
                        // die dummy objekte zur suche wegnehmen da ja nun ein echter ordner angelegt wird
                        DataAdapter.Instance.DataCache.Objects(workspace).RemoveEmptyQueryObjects(true);
                    }

                    // TS 23.04.15
                    if (autosave) typedef = AddDefaultValuesToTypeContainer(typedef, false);

                    // TS 09.10.14 alle relationobjects verarbeiten und diese mitgeben per callback damit sie nachher neu eingelesen werden können
                    if (relationobjects.Count > 0)
                    {
                        DataAdapter.Instance.DataProvider.CreateFoldersInternal(typedef, objecttype, parentobject, typedefrelation, relationobjects, autosave, DataAdapter.Instance.DataCache.Rights.UserPrincipal, new CallbackAction(_CreateFoldersInternal_Done, workspace, selectedobjects, showdialog));
                    }
                    else
                    {
                        DataAdapter.Instance.DataProvider.CreateFoldersInternal(typedef, objecttype, parentobject, typedefrelation, relationobjects, autosave, DataAdapter.Instance.DataCache.Rights.UserPrincipal, new CallbackAction(_CreateFoldersInternal_Done, workspace, null, showdialog));
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        private void _CreateFoldersInternal_Done(CSEnumProfileWorkspace workspace, List<IDocumentOrFolder> selectedobjects, bool showdialog)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (DataAdapter.Instance.DataCache.ResponseStatus.success)
                {
                    // =====================
                    // TS 09.10.14 umbau auf mehrere (wieder rausgenommen, da die indizes nicht für alle gesetzt werden aus dem dialog)
                    DataAdapter.Instance.Processing(workspace).SetSelectedObject(DataAdapter.Instance.DataCache.Objects(workspace).ObjectList[DataAdapter.Instance.DataCache.Objects(workspace).ObjectList.Count - 1].objectId);
                    //List<string> objectidlist = new List<string>();
                    //for (int i = 1; i <= selectedobjects.Count; i++)
                    //{
                    //    objectidlist.Add(DataAdapter.Instance.DataCache.Objects(workspace).ObjectList[DataAdapter.Instance.DataCache.Objects(workspace).ObjectList.Count - i].objectId);
                    //}
                    //DataAdapter.Instance.Processing(workspace).SetSelectedObjects(objectidlist);
                    // =====================

                    // TS 28.01.14 stapel testweise direkt durchwinken
                    if (workspace.ToString().Equals(CSEnumProfileWorkspace.workspace_stapel.ToString()))
                    {
                        //IDocumentOrFolder current = DataAdapter.Instance.DataCache.Objects(workspace).Object_Selected;
                        //current[CSEnumCmisProperties.D_INTERNAL_03] = DateTime.Today.ToShortDateString();
                        //List<cmisObjectType> savelist = new List<cmisObjectType>();
                        //savelist.Add(current.CMISObject);
                        //List<cmisObjectType> parentlist = new List<cmisObjectType>();
                        //parentlist.Add(DataAdapter.Instance.DataCache.Objects(workspace).Root.CMISObject);
                        //DataAdapter.Instance.DataProvider.Save(savelist, parentlist, DataAdapter.Instance.DataCache.Rights.UserPrincipal);
                    }
                    // TS 10.11.14 showdialog dazu
                    else if (showdialog)
                    {
                        SL2C_Client.View.Dialogs.dlgCreateInternalObject child = new View.Dialogs.dlgCreateInternalObject(workspace, selectedobjects);
                        DialogHandler.Show_Dialog(child);
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        private bool _CanCreateFoldersInternal(CSEnumInternalObjectType objecttype, CSEnumProfileWorkspace workspace, int structlevel, List<IDocumentOrFolder> selectedobjects)
        {
            // stapel wird in eigener funktion geprüft
            bool cancreate = DataAdapter.Instance.DataCache.ApplicationFullyInit;
            cancreate = cancreate && DataAdapter.Instance.DataCache.Objects(workspace).Root.canCreateFolder;
            cancreate = cancreate && DataAdapter.Instance.DataCache.Objects(workspace).Root.canCreateObjectLevel(structlevel);

            // zur zeit kann nur ein objekt mitgegeben werden
            // TS 09.10.14 umbau auf mehrere objekte
            if (cancreate && selectedobjects != null && selectedobjects.Count > 0)
            {
                // TS 09.10.14 umbau auf mehrere (wieder rausgenommen, da die indizes nicht für alle gesetzt werden aus dem dialog)
                cancreate = selectedobjects.Count == 1;
                if (cancreate)
                    cancreate = (!selectedobjects[0].isNotCreated && selectedobjects[0].objectId.Length > 0 && selectedobjects[0].isFolder);
                //foreach (IDocumentOrFolder child in selectedobjects)
                //{
                //    if (cancreate) cancreate = (!child.isNotCreated && child.objectId.Length > 0);
                //}
            }
            return cancreate;
        }

        // TS 25.11.13
        //private void _FileOpenDialog(RoutedEventArgs args, CSEnumDocumentCopyTypes copytype, string refid, cmisTypeContainer doctypedefinition, cmisTypeContainer foldertypedefinition,
        //Action<CSEnumDocumentCopyTypes, string, List<CServer.cmisContentStreamType>, cmisTypeContainer, cmisTypeContainer> callback)
        private void _FileOpenDialog(RoutedEventArgs args, CSEnumDocumentCopyTypes copytype, string refid, IDocumentOrFolder parent,
                                        cmisTypeContainer doctypedefinition, cmisTypeContainer foldertypedefinition, bool setaddressflag, Action finalcallback,
                                        Action<List<CServer.cmisContentStreamType>, CSEnumDocumentCopyTypes, string, IDocumentOrFolder, cmisTypeContainer, cmisTypeContainer, bool, Action> callback)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                List<CServer.cmisContentStreamType> contents = new List<cmisContentStreamType>();

                OpenFileDialog dlg = new OpenFileDialog();
                dlg.Filter = "All files (*.*)|*.*";
                dlg.Multiselect = true;
                if (dlg.ShowDialog().Value == true && dlg.Files.Count() > 0)
                {
                    // TS 31.10.12 url dynamisch ermitteln
                    //FileUploadHandler fuh = new FileUploadHandler("http://192.168.100.207:8080/J2C_DMS_CServer/2Charta-DMS/upload",
                    string uri = DataAdapter.Instance.DataProvider.GetProviderAttribs(true, true, true, true, true);
                    //CServer.CServerWSClient proxy = new CServer.CServerWSClient();
                    //System.ServiceModel.Description.ServiceEndpoint endpoint = proxy.Endpoint;
                    //System.ServiceModel.EndpointAddress eaddress = endpoint.Address;
                    //string uri = eaddress.Uri.AbsoluteUri;

                    // TS 25.11.13
                    //FileUploadHandler fuh = new FileUploadHandler(uri + "/upload", dlg.Files.Count(), copytype, refid, doctypedefinition, foldertypedefinition, DataAdapter.Instance.DataCache.Rights.UserPrincipal, callback);
                    FileUploadHandler fuh = new FileUploadHandler(uri + "/upload", dlg.Files.Count(), copytype, refid, parent, doctypedefinition, foldertypedefinition, setaddressflag, DataAdapter.Instance.DataCache.Rights.UserPrincipal, finalcallback, callback);
                    fuh.ProcessFiles(dlg.Files.ToArray());
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        private SaveFileDialog _FileSaveDialog(RoutedEventArgs args, string filename, string filter)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            SaveFileDialog dlg = null;
            try
            {
                dlg = new SaveFileDialog();
                //dlg.Filter = "nur druckbare Dateien (*.pdf)|*.pdf|vollständig (*.zip)|*.zip";

                // TS 18.12.14
                //dlg.Filter = "CSV Dateien (*.csv)|*.csv";
                dlg.Filter = LocalizationMapper.Instance["dlgfilesavefilter_csv"];

                // TS 19.12.13
                if (filter != null && filter.Length > 0)
                    dlg.Filter = filter;
                //string filename = url.Substring(url.LastIndexOf("/") + 1);
                if (filename != null && filename.Length > 0)
                    dlg.DefaultFileName = filename;
                if (dlg.ShowDialog().Value == false)
                    dlg = null;
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
            return dlg;
        }

        /// <summary>
        /// Clears all default-values from the given typdecontainer
        /// </summary>
        /// <param name="typedef"></param>
        /// <returns></returns>
        private void ClearDefaultValuesFromTypeContainer(cmisTypeContainer typedef)
        {
            foreach (cmisPropertyDefinitionType propdeftype in typedef.type.Items)
            {
                switch (propdeftype.propertyType)
                {
                    case enumPropertyType.boolean:
                        ((cmisPropertyBooleanDefinitionType)propdeftype).defaultValue = null;
                        break;

                    case enumPropertyType.datetime:
                        ((cmisPropertyDateTimeDefinitionType)propdeftype).defaultValue = null;
                        break;

                    case enumPropertyType.id:
                        ((cmisPropertyIdDefinitionType)propdeftype).defaultValue = null;
                        break;

                    case enumPropertyType.@decimal:
                        ((cmisPropertyDecimalDefinitionType)propdeftype).defaultValue = null;
                        break;

                    case enumPropertyType.integer:
                        ((cmisPropertyIntegerDefinitionType)propdeftype).defaultValue = null;
                        break;

                    case enumPropertyType.@string:
                        ((cmisPropertyStringDefinitionType)propdeftype).defaultValue = null;
                        break;
                }
            }
        }

        // TS 23.04.15
        private cmisTypeContainer AddDefaultValuesToTypeContainer(cmisTypeContainer typedef, bool isaddress)
        {
            Dictionary<string, string> defaultvalues = null;

            // Clear Defaultvalues
            ClearDefaultValuesFromTypeContainer(typedef);

            // Set new Defaultvalues
            if (isaddress)
                defaultvalues = DataAdapter.Instance.DataCache.Profile.Profile_GetDefaultValues(CSEnumProfileWorkspace.workspace_adressen, typedef.type.id);
            else
                defaultvalues = DataAdapter.Instance.DataCache.Profile.Profile_GetDefaultValues(this.Workspace, typedef.type.id);
            if (defaultvalues != null && defaultvalues.Count > 0)
            {
                foreach (string key in defaultvalues.Keys)
                {
                    foreach (cmisPropertyDefinitionType propdeftype in typedef.type.Items)
                    {
                        if (propdeftype.queryName.Equals(key))
                        {
                            switch (propdeftype.propertyType)
                            {
                                case enumPropertyType.boolean:
                                    cmisPropertyBoolean p0 = new cmisPropertyBoolean();
                                    p0.value = new bool[1];
                                    p0.value[0] = bool.Parse(defaultvalues[key]);
                                    ((cmisPropertyBooleanDefinitionType)propdeftype).defaultValue = p0;
                                    break;

                                case enumPropertyType.datetime:
                                    cmisPropertyDateTime p1 = new cmisPropertyDateTime();
                                    // TS 20.04.17 cmis1.1 patch
                                    //p1.value = new string[1];
                                    p1.value = new DateTime?[1];

                                    string p1value = ValueReplacements.ProcessValueReplacements(defaultvalues[key], DataAdapter.Instance.DataCache.Rights.UserPrincipal);

                                    // TS 27.04.15 falls manuell eingegebenes format dann umwandeln
                                    p1value = DateFormatHelper.GetBasicDateTimeFromUserFormat(p1value);

                                    p1.value[0] = DateFormatHelper.GetCMISDateTime(p1value);
                                    ((cmisPropertyDateTimeDefinitionType)propdeftype).defaultValue = p1;
                                    break;

                                case enumPropertyType.id:
                                    cmisPropertyId p2 = new cmisPropertyId();
                                    p2.value = new string[1];
                                    //p2.value[0] = defaultvalues[key];
                                    p2.value[0] = ValueReplacements.ProcessValueReplacements(defaultvalues[key], DataAdapter.Instance.DataCache.Rights.UserPrincipal);
                                    ((cmisPropertyIdDefinitionType)propdeftype).defaultValue = p2;
                                    break;

                                case enumPropertyType.@decimal:
                                    cmisPropertyDecimal p3 = new cmisPropertyDecimal();
                                    p3.value = new decimal?[1];
                                    p3.value[0] = Decimal.Parse(defaultvalues[key]);
                                    ((cmisPropertyDecimalDefinitionType)propdeftype).defaultValue = p3;
                                    break;

                                case enumPropertyType.integer:
                                    cmisPropertyInteger p4 = new cmisPropertyInteger();
                                    p4.value = new string[1];
                                    p4.value[0] = defaultvalues[key];
                                    ((cmisPropertyIntegerDefinitionType)propdeftype).defaultValue = p4;
                                    break;

                                case enumPropertyType.@string:
                                    cmisPropertyString p5 = new cmisPropertyString();
                                    p5.value = new string[1];
                                    //p5.value[0] = defaultvalues[key];
                                    p5.value[0] = ValueReplacements.ProcessValueReplacements(defaultvalues[key], DataAdapter.Instance.DataCache.Rights.UserPrincipal);
                                    ((cmisPropertyStringDefinitionType)propdeftype).defaultValue = p5;
                                    break;
                            }
                            break;
                        }
                    }
                }
            }
            return typedef;
        }

        #endregion diverse privates

        #region _Query (privates)

        /// <summary>
        /// querymode 0 = neue suche / 1 = nächste treffer / 2 = vorherige treffer
        /// </summary>
        /// <param name="querymode"></param>
        private void _Query(QueryMode querymode, CSEnumProfileWorkspace targetworkspace) { _Query(querymode, "", targetworkspace); }

        private void _Query(QueryMode querymode, CSEnumProfileWorkspace targetworkspace, CallbackAction callback, CallbackAction notDoneCallback)
        {
            _Query(querymode, "", targetworkspace, callback, notDoneCallback);
        }

        private void _Query(QueryMode querymode, string in_folder, CSEnumProfileWorkspace targetworkspace)
        {
            _Query(querymode, in_folder, targetworkspace, null, null);
        }

        /// <summary>
        /// querymode 0 = neue suche / 1 = nächste treffer / 2 = vorherige treffer
        /// in_folder = suche nur in diesem ordner
        /// </summary>
        /// <param name="querymode"></param>
        /// <param name="in_folder"></param>
        /// <param name="targetworkspace"></param>
        /// <param name="finalCallback"></param>
        private void _Query(QueryMode querymode, string in_folder, CSEnumProfileWorkspace targetworkspace, CallbackAction finalCallbackOnReadMore, CallbackAction notDoneCallback)
        {
            // je nach suchtyp den cache vorher leeren
            // TS 27.03.15
            // if (querymode != QueryMode.QueryNew)
            if (querymode != QueryMode.QueryNew && querymode != QueryMode.QueryMore)
            {
                // TS 25.05.16
                //ClearCache(false, new CallbackAction(_Query_Proceed, querymode, in_folder, targetworkspace));
                DataAdapter.Instance.Processing(targetworkspace).ClearCache(false, new CallbackAction(_Query_Proceed, querymode, in_folder, targetworkspace));
            }
            else
                _Query_Proceed(querymode, in_folder, targetworkspace, finalCallbackOnReadMore, notDoneCallback);
        }

        // TS 20.03.14 weiterer eingang zum direkten mitgeben von suchbedingungen (wird bisher nur verwendet aus EDesktop.QeryGesamt)
        private void _QueryGiven(QueryMode querymode, List<string> displayproperties, Dictionary<string, List<CSQueryToken>> repositoryquerytokens, Dictionary<string, List<CSOrderToken>> repositoryordertokens,
            bool autotruncate, bool getimagelist, bool getfulltextpreview, bool getparents, bool usesearchengine, CSEnumProfileWorkspace targetworkspace, CallbackAction callback)
        {
            // je nach suchtyp den cache vorher leeren
            // TS 27.03.15
            // if (querymode != QueryMode.QueryNew)
            if (querymode != QueryMode.QueryNew && querymode != QueryMode.QueryMore)
                ClearCache(false, new CallbackAction(_QueryGiven_Proceed, querymode, displayproperties, repositoryquerytokens, repositoryordertokens, autotruncate, getimagelist, getfulltextpreview, getparents, usesearchengine, targetworkspace));
            else
                _QueryGiven_Proceed(querymode, displayproperties, repositoryquerytokens, repositoryordertokens, autotruncate, getimagelist, getfulltextpreview, getparents, usesearchengine, targetworkspace, callback);
        }

        private void _Query_Proceed(QueryMode querymode, string in_folder, CSEnumProfileWorkspace targetworkspace)
        {
            _Query_Proceed(querymode, in_folder, targetworkspace, null, null);
        }
        private void _Query_Proceed(QueryMode querymode, string in_folder, CSEnumProfileWorkspace targetworkspace, CallbackAction finalCallback, CallbackAction notDoneCallback)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                DataAdapter.Instance.DataCache.Info.IsQueryOverMatchlist = false;
                if (this.Workspace.ToString().Equals(CSEnumProfileWorkspace.workspace_deleted.ToString()))
                    _QueryDeleted(querymode, targetworkspace, finalCallback);
                else if (this.Workspace.ToString().Equals(CSEnumProfileWorkspace.workspace_lastedited.ToString()))
                    _QueryLastEdited(querymode, targetworkspace, finalCallback);
                else
                {
                    // vorbelegungen
                    List<CSQueryProperties> querypropertieslist = new List<CSQueryProperties>();
                    DataAdapter.Instance.DataCache.Objects(Workspace).QueryCollectionURL = new Uri("", UriKind.RelativeOrAbsolute);
                    DataAdapter.Instance.DataCache.Objects(Workspace).QueryCollectionXML = "";
                    // TODO: Umleitung Suche von default nach overall workspace verbessern (nicht hier hart verdrahtet)

                    // TS 10.02.17 die adressen nur wenn nicht im eigenen workspace
                    // HIER EINBLENDEN fuer Adressen in eigenem workspace
                    //bool isdefaultworkspace = targetworkspace.ToString().Equals(CSEnumProfileWorkspace.workspace_default.ToString());
                    // TS 12.07.17
                    bool isdefaultworkspace = targetworkspace.ToString().Equals(CSEnumProfileWorkspace.workspace_default.ToString()) 
                        || targetworkspace.ToString().Equals(CSEnumProfileWorkspace.workspace_adressen.ToString());

                    bool shiftworkspaceoverall = isdefaultworkspace;

                    IProcessing overallprocessing = this;
                    if (shiftworkspaceoverall) overallprocessing = DataAdapter.Instance.Processing(CSEnumProfileWorkspace.workspace_searchoverall);
                    // TS 04.06.18 nach unten
                    //bool getQueryCountDetails = isdefaultworkspace; // TODO: Get this from an profiles' option some day
                    bool autotruncate = DataAdapter.Instance.DataCache.Profile.Option_GetValueBoolean(CSEnumOptions.queryautotrunc);
                    bool getfulltextpreview = true;
                    // bilderliste nur im default workspace
                    bool getimagelist = isdefaultworkspace && DataAdapter.Instance.DataCache.Profile.Option_GetValueBoolean(CSEnumOptions.queryshowimagelist);
                    bool usesearchengine = false;
                    bool getparents = true;

                    // TS 04.06.18 die meldung je nach schaltern von der queryCount liefern lassen weil da die gesamtsumme kommt und nicht nur die ersten 50
                    bool getQueryCountDetails = isdefaultworkspace; // TODO: Get this from an profiles' option some day
                    // der hier war true
                    //bool showquerycount = true; 
                    bool showQueryCountRegular = true;
                    bool showQueryCountOverall = false;

                    // nur wenn ueberhaupt eine ausgabe gewuenscht ist
                    if (showQueryCountRegular)
                    {
                        // wenn details aktiv sind dann kann eine summierte anzahl ermittelt werden
                        if (getQueryCountDetails)
                        {
                            showQueryCountOverall = true;
                            showQueryCountRegular = false;
                        }
                    }

                    int querylistsize = int.Parse(DataAdapter.Instance.DataCache.Profile.Option_GetValue(CSEnumOptions.querylistsize));                    
                    
                    if(querymode == QueryMode.QueryMore)
                    {
                        querylistsize = int.Parse(DataAdapter.Instance.DataCache.Profile.Option_GetValue(CSEnumOptions.querytreesize));
                    }

                    // TS 20.03.17 fuer adressen selektion die anzahl hochsetzen und das autotrunc abschalten
                    if (Workspace == CSEnumProfileWorkspace.workspace_adressenselektion)
                    {
                        querylistsize = Constants.QUERY_ADRSELEKTION_MAXITEMS;
                        autotruncate = false;
                    }

                    int queryskipcount = 0;
                    QueryMode tmpquerymode = querymode == QueryMode.QueryMore ? QueryMode.QueryNext : querymode;

                    bool hasvalues = tmpquerymode == QueryMode.QueryRefresh && DataAdapter.Instance.DataCache.Objects(this.Workspace).Query_LastQueryValues.Count > 0;
                    // TS 14.11.16 vereinfachung wegen besserer uebersicht
                    bool isnewquery = tmpquerymode != QueryMode.QueryPrev && tmpquerymode != QueryMode.QueryNext && (tmpquerymode != QueryMode.QueryRefresh || !hasvalues);
                    // ******************************************************************************************************************************************************
                    if (!isnewquery)
                    {
                        // TS 20.03.14 suchbedingungen aus datacache_objects holen
                        querypropertieslist = DataAdapter.Instance.DataCache.Objects(this.Workspace).Query_LastQueryValues;
                        // TS 15.10.15 die sortierung live neu ermitteln, klappt aber erstmal nur für einen satz von bedingungen sonst müßte ich zu tief einsteigen hier
                        if (querypropertieslist.Count == 1)
                        {
                            List<CServer.CSOrderToken> orderby = new List<CSOrderToken>();
                            _Query_GetSortDataFilter(ref orderby, true);
                            if (orderby.Count > 0)
                            {
                                CSOrderToken[] aorderby = new CSOrderToken[orderby.Count];
                                for (int i = 0; i < orderby.Count; i++) { aorderby.SetValue(orderby[i], i); }
                                querypropertieslist[0].orderby = aorderby;
                            }
                        }
                        _Query_GetSkipCount(tmpquerymode, querypropertieslist);
                        // TS 16.08.16 fix für gespeicherte suchen über SE
                        foreach (CSQueryProperties props in querypropertieslist)
                        {
                            foreach (CSQueryToken token in props.searchconditions)
                            {
                                if (token.propertyname.Equals("CONTAINS"))
                                {
                                    usesearchengine = true;
                                    break;
                                }
                            }
                        }
                    }
                    // ******************************************************************************************************************************************************
                    else if (isnewquery)
                    {
                        // je nach suchtyp, entweder suchdaten ermitteln oder aus den gespeicherten holen
                        Dictionary<string, string> queryvalues = DataAdapter.Instance.DataCache.Objects(Workspace).CollectQueryValues(tmpquerymode);
                        // zu durchsuchende repositories, werden zuerst über datenfilter gesetzt wenn vorhanden, sonst das standard-reptry vom workspace
                        List<string> searchrepositories = new List<string>();
                        // baut querytokens aus den abgeholten werten, schalter ob fulltext oder nicht
                        List<CServer.CSQueryToken> querytokens = new List<CSQueryToken>();

                        // fulltext tokens bauen und auswerten
                        List<CServer.CSQueryToken> querytokensfulltext = _Query_CreateQueryTokens(queryvalues, true);
                        if (querytokensfulltext.Count > 0)
                        {
                            // hier umstellen auf neue suchlogik sobald verfügbar
                            bool searchusenewlogic = false;
                            int fulltextlevel = 1;
                            try { fulltextlevel = DataAdapter.Instance.DataCache.Profile.UserProfile.fulltextlevel; }
                            catch (Exception) { }
                            searchusenewlogic = fulltextlevel == 2;
                            _Query_PrepareFulltextQuery(searchusenewlogic, ref querytokens, ref querytokensfulltext, ref in_folder, ref usesearchengine, ref searchrepositories, ref targetworkspace);
                        }
                        else
                        {
                            // holt die filter aus dem aktuellen workspace (nicht overall)
                            _Query_AddDataFilter(ref searchrepositories, ref queryvalues);
                            querytokens = _Query_CreateQueryTokens(queryvalues, false);

                            // TS 26.08.16 übergreifende suche, aber nur für echten aktenplan
                            // TS 12.07.17
                            // if (isdefaultworkspace && DataCache_RootNodes.RealRootNodesAvail)
                            if (DataCache_RootNodes.RealRootNodesAvail)
                            {
                                if (isdefaultworkspace)
                                {
                                    _Query_GetSearchRepositories(null, ref searchrepositories);
                                }
                                    // TS 12.07.17
                                else if (targetworkspace.ToString().Equals(CSEnumProfileWorkspace.workspace_adressenselektion.ToString()))
                                {
                                    _Query_GetSearchRepositories(CSEnumProfileWorkspace.workspace_default, null, ref searchrepositories);
                                }
                            }
                            if (searchrepositories.Count == 0)
                                searchrepositories.Add(DataAdapter.Instance.DataCache.Repository(Workspace).RepositoryInfo.repositoryId);
                        }
                        // TS 22.01.14 nur wenn nicht volltext
                        // TS 02.11.17 in_folder wird nicht mehr mit "and parentfolders:xyz" übergeben sondern mit in_folder(xyz) und daher hier auch behandelt
                        // if (querytokensfulltext.Count == 0 && in_folder != null && in_folder.Length > 0)
                        if (in_folder != null && in_folder.Length > 0)
                        {
                            CSQueryToken token = new CSQueryToken();

                            // TS 03.01.18 hier: änderung auf IN_TREE
                            //token.propertyname = "IN_FOLDER";
                            // insgesamt 4 Stellen im SL Code
                            // 1. Query_EDP_Gesamt: Suche in Posteingang => bleibt IN_FOLDER
                            // 2. Query_EDP_Gesamt: Suche in Maileingang => bleibt IN_FOLDER
                            // 3. Query_Depending:                       => wird   IN_TREE
                            // 4. _Query_Proceed (hier):                 => wird   IN_TREE
                            token.propertyname = "IN_TREE";

                            token.propertyvalue = in_folder;
                            token.propertytype = enumPropertyType.@string;
                            querytokens.Add(token);
                        }
                        // sortierung, abholen von orderby presets
                        List<CServer.CSOrderToken> orderby = new List<CSOrderToken>();
                        _Query_GetSortDataFilter(ref orderby, true, overallprocessing);

                        if (orderby.Count == 0)
                        {
                            CSOrderToken firstorderby = new CSOrderToken();
                            firstorderby.propertyname = "cmis:objectId";
                            firstorderby.orderby = CSEnumOrderBy.asc;
                            orderby.Add(firstorderby);
                        }
                        // TS 20.01.16 nicht mehr mit "*" suchen sondern nur noch die konkret ausgewählten indizes aus trefferliste und treeview verwenden
                        List<string> displayproperties = _Query_GetDisplayProperties(true, true, overallprocessing);
                        foreach (string searchrepository in searchrepositories)
                        {
                            CSQueryProperties queryproperties = _QueryCreateQueryPropertiesFromValues(searchrepository, displayproperties, querytokens, orderby, autotruncate, queryskipcount, usesearchengine);

                            // TS 03.01.18
                            if (usesearchengine)
                            {
                                try
                                {
                                    string fulltextlevel = DataAdapter.Instance.DataCache.Profile.Option_GetValue(CSEnumOptions.queryfulltextlevel);
                                    if (fulltextlevel != null && fulltextlevel.Length > 0)
                                    {
                                        queryproperties.queryfulltextlevel = fulltextlevel;
                                    }
                                }
                                catch (Exception) { }
                            }

                            querypropertieslist.Add(queryproperties);
                        }

                        // suchbedingungen in datacache_objects speichern
                        // TS 05.08.16 die prüfung auf volltext rausgenommen da diese Werte für Refresh verwendet werden 
                        // und ich vermute es ist wegen der suchen in EDesktop, daher Einschränken auf workspace_default
                        //if (querytokensfulltext.Count == 0 && (!HasFixedQuery || DataAdapter.Instance.DataCache.Objects(this.Workspace).Query_LastQueryValues.Count == 0))
                        if ((querytokensfulltext.Count == 0 || isdefaultworkspace) && (!HasFixedQuery || DataAdapter.Instance.DataCache.Objects(Workspace).Query_LastQueryValues.Count == 0))
                            DataAdapter.Instance.DataCache.Objects(this.Workspace).Query_LastQueryValues = querypropertieslist;

                        // muss das nicht in die klammer ?? oder kann es womöglich ganz weg ??
                        if (isdefaultworkspace)
                        {
                            DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_searchoverall).Query_LastQueryValues = querypropertieslist;
                        }
                    }
                    // ******************************************************************************************************************************************************
                    
                    // Get an optional List of last queryproperties
                    List<CSQueryProperties> lastQueryList = new List<CSQueryProperties>();
                    lastQueryList = CheckAndApplyPossibleQueryChronicles(querypropertieslist);

                    // ******************************************************************************************************************************************************
                    // TS 01.07.14 wenn keine suchbedingungen dann auch keine suche
                    bool isvalidquery = querypropertieslist.Count() > 0 && querypropertieslist[0].searchconditions.Count() > 0;
                    if (!isvalidquery)
                    {
                        // warnung ausgeben
                        DisplayWarnMessage(LocalizationMapper.Instance["msg_query_emptyquery"]);
                        // aktualisieren damit die buttons zurückschalten
                        DataAdapter.Instance.InformObservers();
                        // evtl. vorhandenes Callback auslösen
                        if (notDoneCallback != null)
                        {
                            notDoneCallback.Invoke();
                            notDoneCallback = null;
                        }
                    }
                    // ******************************************************************************************************************************************************
                    // suche in bpm
                    if (isvalidquery && querypropertieslist[0].searchrepository.StartsWith(Statics.Constants.REPOSITORY_WORKFLOW.ToString() + "_"))
                    {
                        DataAdapter.Instance.DataProvider.BPM_Query(querypropertieslist, getparents, getfulltextpreview, getimagelist, usesearchengine, querylistsize, targetworkspace,
                        DataAdapter.Instance.DataCache.Rights.UserPrincipal, new CallbackAction(Query_Done, querymode, finalCallback));
                    }
                    // normale suche
                    else if (isvalidquery)
                    {
                        if (isdefaultworkspace && querymode == QueryMode.QueryNew)
                        {
                            ClearCache();
                            QuerySetLast();
                        }
                        if (isdefaultworkspace && shiftworkspaceoverall) targetworkspace = CSEnumProfileWorkspace.workspace_searchoverall;

                        //bool showQueryCountRegular = true;
                        //bool showQueryCountOverall = false;

                        // TS 04.06.18
                        //DataAdapter.Instance.DataProvider.Query(querypropertieslist, lastQueryList, getparents, getfulltextpreview, getimagelist, usesearchengine, querylistsize, targetworkspace, showquerycount, 
                        //DataAdapter.Instance.DataCache.Rights.UserPrincipal, new CallbackAction(overallprocessing.Query_Done, querymode, finalCallback));
                        DataAdapter.Instance.DataProvider.Query(querypropertieslist, lastQueryList, getparents, getfulltextpreview, getimagelist, usesearchengine, querylistsize, targetworkspace, showQueryCountRegular,
                        DataAdapter.Instance.DataCache.Rights.UserPrincipal, new CallbackAction(overallprocessing.Query_Done, querymode, finalCallback));

                        if (getQueryCountDetails)
                        {
                            // TS 04.06.18
                            //DataAdapter.Instance.DataProvider.QueryCount(querypropertieslist, lastQueryList, getparents, usesearchengine, true, true, targetworkspace,
                            //DataAdapter.Instance.DataCache.Rights.UserPrincipal, new CallbackAction(RefreshQueryCount));
                            DataAdapter.Instance.DataProvider.QueryCount(querypropertieslist, lastQueryList, getparents, usesearchengine, true, true, showQueryCountOverall, targetworkspace, 
                            DataAdapter.Instance.DataCache.Rights.UserPrincipal, new CallbackAction(RefreshQueryCount));
                        }                        
                    }
                    // ******************************************************************************************************************************************************
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        private void _Query_PrepareFulltextQuery(bool searchusenewlogic, ref List<CServer.CSQueryToken> querytokens, ref List<CServer.CSQueryToken> querytokensfulltext, ref string in_folder, ref bool usesearchengine, ref List<string> searchrepositories, ref CSEnumProfileWorkspace targetworkspace)
        {
            CSQueryToken fulltexttoken = new CSQueryToken();
            fulltexttoken.propertyname = "CONTAINS";
            fulltexttoken.propertyvalue = "";
            fulltexttoken.propertytype = enumPropertyType.@string;
            // TS 26.11.12 zur zeit kann nur ein Volltext Objekttyp zur Suche verwendet werden
            fulltexttoken.propertyreptypeid = querytokensfulltext.First().propertyreptypeid;
            foreach (CServer.CSQueryToken temptoken in querytokensfulltext)
            {
                if (fulltexttoken.propertyvalue.Length > 0)
                {
                    fulltexttoken.propertyvalue = fulltexttoken.propertyvalue + Constants.QUERY_OPERATOR_OR;
                }
                // TS 18.03.14 wg. SearchEngine nicht mehr alle fulltext properties abfragen
                // sondern nur noch fix inhalt und selbst der wird später entfernt und nur noch das reine suchvalue übergeben
                // fulltexttoken.propertyvalue = fulltexttoken.propertyvalue + temptoken.propertyname + ":" + temptoken.propertyvalue;
                if (temptoken.propertyname.Equals(Constants.FULLTEXT_SEARCHENGINE))
                    fulltexttoken.propertyvalue = fulltexttoken.propertyvalue + temptoken.propertyvalue;
                else
                    fulltexttoken.propertyvalue = fulltexttoken.propertyvalue + temptoken.propertyname + ":" + temptoken.propertyvalue;
            }
            // TS 02.11.17 in_folder wird nicht mehr mit "and parentfolders:xyz" übergeben sondern mit in_folder(xyz) und daher weiter oben behandelt
            //// TS 24.05.16 die 0 kommt jetzt mit wenn innerhalb eines repositories gesucht werden soll
            //// if (in_folder != null && in_folder.Length > 0)
            //if (in_folder != null && in_folder.Length > 1 && !in_folder.Equals(DataAdapter.Instance.DataCache.Objects(targetworkspace).Root.objectId))
            //{
            //    if (fulltexttoken.propertyvalue.Length > 0)
            //    {
            //        fulltexttoken.propertyvalue = "(" + fulltexttoken.propertyvalue + ")" + " AND ";
            //    }
            //    // das "O" abschneiden
            //    in_folder = in_folder.Substring(1);
            //    fulltexttoken.propertyvalue = fulltexttoken.propertyvalue + "parentfolders" + ":" + in_folder;
            //}
            // TS 09.02.15 für die suche nach aufgaben und terminen eine textkonstante mitgeben da diese ja leider nicht an einem parent hängen
            if (this.Workspace.ToString().Equals(CSEnumProfileWorkspace.workspace_aufgabe.ToString()))
            {
                fulltexttoken.propertyvalue = Statics.Constants.FULLTEXT_AUFGABEKEY + " " + fulltexttoken.propertyvalue;
            }
            else if (this.Workspace.ToString().Equals(CSEnumProfileWorkspace.workspace_termin.ToString()))
            {
                fulltexttoken.propertyvalue = Statics.Constants.FULLTEXT_TERMINKEY + " " + fulltexttoken.propertyvalue;
            }

            querytokens.Add(fulltexttoken);
            usesearchengine = true;

            // TS 11.06.15 sonderbehandlung
            if (targetworkspace == CSEnumProfileWorkspace.workspace_adressen)
            {
                // TS 17.07.17 raus da so nicht mehr verwendet
                //searchrepositories.Add(DataAdapter.Instance.DataCache.Objects(targetworkspace).Root.RepositoryId);
                if (querytokens.Count > 0)
                {
                    querytokens[querytokens.Count - 1].propertyvalue = querytokens[querytokens.Count - 1].propertyvalue + Constants.QUERY_OPERATOR_AND + Constants.FULLTEXT_ADRESSFLAG;
                }
                // TS 01.03.17 raus fuer adressen in eigenem workspace
                // targetworkspace = CSEnumProfileWorkspace.workspace_default;
            }
            // TS 17.07.17 das else auch raus da so nicht mehr verwendet
            //else
            //{
            if (in_folder != null && in_folder.Length > 0)
            {
                searchrepositories.Add(DataAdapter.Instance.DataCache.Objects(this.Workspace).Root.RepositoryId);
            }
            // TS 24.05.16 nur übergreifend suchen wenn nicht in_folder angegeben
            // if (in_folder == null || in_folder.Length == 0)
            else
            {
                // TS 17.07.17
                // List<SearchRepositoryToken> searchreps = DataAdapter.Instance.DataCache.Objects(this.Workspace).QuerySE_SearchRepositories;
                // ********************************************************************************
                bool isdefaultworkspace = targetworkspace.ToString().Equals(CSEnumProfileWorkspace.workspace_default.ToString())
                    || targetworkspace.ToString().Equals(CSEnumProfileWorkspace.workspace_adressen.ToString());

                if (DataCache_RootNodes.RealRootNodesAvail)
                {
                    if (isdefaultworkspace)
                    {
                        _Query_GetSearchRepositories(null, ref searchrepositories);
                    }
                    else if (targetworkspace.ToString().Equals(CSEnumProfileWorkspace.workspace_adressenselektion.ToString()))
                    {
                        _Query_GetSearchRepositories(CSEnumProfileWorkspace.workspace_default, null, ref searchrepositories);
                    }
                }
                if (searchrepositories.Count == 0)
                    searchrepositories.Add(DataAdapter.Instance.DataCache.Repository(Workspace).RepositoryInfo.repositoryId);

                // ********************************************************************************

                // alte logik mit alter searchengine
                if (!searchusenewlogic)
                {
                    // TS 17.07.17
                    List<SearchRepositoryToken> searchreps = DataAdapter.Instance.DataCache.Objects(this.Workspace).QuerySE_SearchRepositories;
                    foreach (SearchRepositoryToken token in searchreps)
                    {
                        if (token.IsSelected)
                        {
                            // nicht alle aktenplanknoten übergeben, das klappt nicht sonderlich gut, dann lieber nur das aktuelle archiv falls es enthalten ist
                            //foreach (string repid in token.RepositoryIds)
                            //{
                            //    searchrepositories.Add(repid);
                            //}
                            if (token.RepositoryIds.Count > 1)
                            {
                                if (token.RepositoryIds.Contains(DataAdapter.Instance.DataCache.Rights.ChoosenRepository.repositoryId) 
                                    && !searchrepositories.Contains(DataAdapter.Instance.DataCache.Rights.ChoosenRepository.repositoryId))
                                {
                                    searchrepositories.Add(DataAdapter.Instance.DataCache.Rights.ChoosenRepository.repositoryId);
                                }
                            }
                            else if(!searchrepositories.Contains(token.RepositoryIds.First()))
                            {
                                searchrepositories.Add(token.RepositoryIds.First());
                            }
                                
                        }
                    }
                }
                else if (searchusenewlogic)
                {
                    // TS 17.07.17 das muesste bereits erledigt sein
                    //_Query_GetSearchRepositories(searchreps, ref searchrepositories);
                }
            }
            //}
        }

        private void _Query_GetSearchRepositories(List<SearchRepositoryToken> searchreps, ref List<string> searchrepositories)
        {
            _Query_GetSearchRepositories(this.Workspace, searchreps, ref searchrepositories);
        }

        private void _Query_GetSearchRepositories(CSEnumProfileWorkspace workspace, List<SearchRepositoryToken> searchreps, ref List<string> searchrepositories)
        {
            if (searchreps == null) searchreps = DataAdapter.Instance.DataCache.Objects(workspace).QuerySE_SearchRepositories;

            // *******************************************************************************************
            // TS 22.06.17 umbau mal wieder, siehe bei QuerySE_SearchRepositories
            // das erste token ist immer "nur aktuellen knoten durchsuchen" oder so
            // wenn des NICHT gesetzt ist wird uebergreifend gesucht
            //searchrepositories.Add("0" + "_" + DataAdapter.Instance.DataCache.Rights.ChoosenMandant.id);
            if (searchreps.Count > 0)
            {
                bool first = true;
                bool overall = false;
                foreach (SearchRepositoryToken token in searchreps)
                {
                    if (first)
                    {
                        if (token.IsSelected)
                        {
                            searchrepositories.Add(token.RepositoryIds.First());
                        }
                        else
                        {
                            // uebergreifend
                            searchrepositories.Add("0" + "_" + DataAdapter.Instance.DataCache.Rights.ChoosenMandant.id);
                            overall = true;
                        }
                        first = false;
                    }
                    else
                    {
                        // alle anderen, post, mail
                        // wenn oben overall gesetzt und die hier NICHT gewählt dann mit minus rausrechnen
                        // sonst die hier reinrechnen
                        if (overall && !token.IsSelected)
                        {
                            searchrepositories.Add("-" + token.RepositoryIds.First());
                        }
                        else if (!overall && token.IsSelected)
                        {
                            searchrepositories.Add(token.RepositoryIds.First());
                        }
                    }
                }
            }
            else
            {
                // nichts gefunden dann uebergreifend suchen, besser als garnicht
                searchrepositories.Add("0" + "_" + DataAdapter.Instance.DataCache.Rights.ChoosenMandant.id);
            }
            // *******************************************************************************************

            //// ==============================================================================================================================================================
            //// TS 25.05.16 hier muss jetzt eine neue logik hinein
            //// falls alle der ausgewählten SearchRepositoryTokens nur genau ein repository enthalten
            ////      => dann werden diese einfach angegeben so wie vorher auch
            ////
            //// falls aber genau eines der ausgewählten SearchRepositoryTokens mehr als ein repository enthält (z.b. bei Aktenplan)
            ////      => dann wird dieses mit "0" an die searchengine übergeben
            ////      => und alle weiteren ausgewählten SearchRepositoryTokens werden weggelassen (die 0 von oben sucht ja übergreifend)
            ////      => und alle weiteren NICHT ausgewählten SearchRepositoryTokens werden negativ angegeben (*-1) damit diese bei der übergreifenden suche weggelassen werden
            ////
            //// falls gar mehrere ausgewählte SearchRepositoryTokens mehr als ein repository enthalten (z.b. bei Aktenplan mit unterschiedlichen Anwendungen)
            ////      => dann wird geprüft, ob die menge der zu durchsuchenden repositories größer ist als die menge der NICHT zu durchsuchenden
            ////      => wenn ja, werden die zu durchsuchenden als "0" angegeben und die NICHT zu durchsuchenden negativ (*-1) um sie wegzulassen
            ////      => wenn nein, werden die zu durchsuchenden alle einzeln angegeben
            //// ==============================================================================================================================================================

            //// prüfen ob eins oder mehrere tokens mehr als ein repository haben
            //int anzahlfound = 0;
            //foreach (SearchRepositoryToken token in searchreps)
            //{
            //    if (token.IsSelected && token.RepositoryIds.Count > 1)
            //        anzahlfound++;
            //}

            //// 1. alle haben nur ein repositiory
            //if (anzahlfound == 0)
            //{
            //    foreach (SearchRepositoryToken token in searchreps)
            //    {
            //        if (token.IsSelected)
            //        {
            //            searchrepositories.Add(token.RepositoryIds.First());
            //        }
            //    }
            //}

            //// 2. genau eines hat mehr als ein repository
            //else if (anzahlfound == 1)
            //{
            //    // TS 06.07.16 erstmal alle id sammeln damit nicht negative angegeben werden die in den positiven enthalten sind
            //    // konkretes beispiel aktenplan: die knoten "schriftgutverwaltung" und "0 Allgemeine Verwaltung" sind angewählt
            //    // wobei das repository von "0 Allgemeine Verwaltung" in der menge von "schriftgutverwaltung" enthalten ist
            //    // dann macht es keinen sinn die "0 Allgemeine Verwaltung" als negativ anzugeben denn die soll ja auch durchsucht werden
            //    List<string> positives = new List<string>();
            //    foreach (SearchRepositoryToken token in searchreps)
            //    {
            //        if (token.IsSelected && token.RepositoryIds.Count > 1)
            //        {
            //            positives.AddRange(token.RepositoryIds);
            //        }
            //    }
            //    foreach (SearchRepositoryToken token in searchreps)
            //    {
            //        if (token.IsSelected && token.RepositoryIds.Count > 1)
            //        {
            //            searchrepositories.Add("0" + "_" + DataAdapter.Instance.DataCache.Rights.ChoosenMandant.id);
            //        }
            //        else if (token.IsSelected && token.RepositoryIds.Count == 1)
            //        {
            //            // die werden weggelassen
            //        }
            //        else if (!token.IsSelected && token.IsEnabled)
            //        {
            //            // alle anderen negativ angeben
            //            foreach (string repid in token.RepositoryIds)
            //            {
            //                // TS 06.07.16 nicht rausnehmen wenn in derr positivliste (siehe oben)
            //                if (!positives.Contains(repid))
            //                {
            //                    searchrepositories.Add("-" + repid);
            //                }
            //            }
            //        }
            //    }
            //}

            //// 3. mehrere haben mehr als ein repository
            //else if (anzahlfound > 1)
            //{
            //    int anzahlsearch = 0;
            //    int anzahlnotsearch = 0;
            //    foreach (SearchRepositoryToken token in searchreps)
            //    {
            //        if (token.IsSelected)
            //        {
            //            anzahlsearch = anzahlsearch + token.RepositoryIds.Count();
            //        }
            //        else if (!token.IsSelected && token.IsEnabled)
            //        {
            //            anzahlnotsearch = anzahlnotsearch + token.RepositoryIds.Count();
            //        }
            //    }
            //    if (anzahlsearch > anzahlnotsearch)
            //    {
            //        // TS 06.07.16 erstmal alle id sammeln damit nicht negative angegeben werden die in den positiven enthalten sind (siehe oben)
            //        List<string> positives = new List<string>();
            //        foreach (SearchRepositoryToken token in searchreps)
            //        {
            //            if (token.IsSelected && token.RepositoryIds.Count > 1)
            //            {
            //                positives.AddRange(token.RepositoryIds);
            //            }
            //        }

            //        searchrepositories.Add("0" + "_" + DataAdapter.Instance.DataCache.Rights.ChoosenMandant.id);

            //        foreach (SearchRepositoryToken token in searchreps)
            //        {
            //            if (!token.IsSelected && token.IsEnabled)
            //            {
            //                foreach (string repid in token.RepositoryIds)
            //                {
            //                    // TS 06.07.16 nicht rausnehmen wenn in derr positivliste (siehe oben)
            //                    if (!positives.Contains(repid))
            //                    {
            //                        searchrepositories.Add("-" + repid);
            //                    }
            //                }
            //            }
            //        }
            //    }
            //    else
            //    {
            //        foreach (SearchRepositoryToken token in searchreps)
            //        {
            //            if (token.IsSelected)
            //            {
            //                foreach (string repid in token.RepositoryIds)
            //                {
            //                    searchrepositories.Add(repid);
            //                }
            //            }
            //        }
            //    }
            //}
            //// ==============================================================================================================================================================
        }

        // TS 20.03.14 weiterer eingang zum direkten mitgeben von suchbedingungen (wird bisher nur verwendet aus EDesktop.QeryGesamt)
        // TS 22.01.16 alle parameter mitgeben
        // private void _Query_Proceed(QueryMode querymode, Dictionary<string, List<CSQueryToken>> repositoryquerytokens, Dictionary<string, List<CSOrderToken>> repositoryordertokens, CSEnumProfileWorkspace targetworkspace)
        private void _QueryGiven_Proceed(QueryMode querymode, List<string> displayproperties, Dictionary<string, List<CSQueryToken>> repositoryquerytokens, Dictionary<string, List<CSOrderToken>> repositoryordertokens,
            bool autotruncate, bool getimagelist, bool getfulltextpreview, bool getparents, bool usesearchengine, CSEnumProfileWorkspace targetworkspace)
        {
            _QueryGiven_Proceed(querymode, displayproperties, repositoryquerytokens, repositoryordertokens, autotruncate, getimagelist, getfulltextpreview, getparents, usesearchengine, targetworkspace, null);
        }

        private void _QueryGiven_Proceed(QueryMode querymode, List<string> displayproperties, Dictionary<string, List<CSQueryToken>> repositoryquerytokens, Dictionary<string, List<CSOrderToken>> repositoryordertokens,
            bool autotruncate, bool getimagelist, bool getfulltextpreview, bool getparents, bool usesearchengine, CSEnumProfileWorkspace targetworkspace, CallbackAction callback)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                // TS 26.09.14 meldung
                // TS 14.10.15 raus erstmal da durch die logik in datacache zu oft diese meldung kam, nach umbau dort dann aber garnicht mehr, also raus
                //DataAdapter.Instance.DataCache.ResponseStatus.localizedmessage = LocalizationMapper.Instance["msg_query_start"];

                if(targetworkspace == CSEnumProfileWorkspace.workspace_default)
                {
                    targetworkspace = CSEnumProfileWorkspace.workspace_searchoverall;
                }

                // vorbelegungen
                List<CSQueryProperties> querypropertieslist = new List<CSQueryProperties>();
                DataAdapter.Instance.DataCache.Objects(Workspace).QueryCollectionURL = new Uri("", UriKind.RelativeOrAbsolute);
                DataAdapter.Instance.DataCache.Objects(Workspace).QueryCollectionXML = "";

                //bool autotruncate = DataAdapter.Instance.DataCache.Profile.Option_GetValueBoolean(CSEnumOptions.queryautotrunc);

                //// TS 23.05.14 bilderliste nicht immer lesen sondern nur im default workspace
                ////bool getimagelist = DataAdapter.Instance.DataCache.Profile.Option_GetValueBoolean(CSEnumOptions.showimagelist);
                //bool getimagelist = false;
                //if (targetworkspace.ToString().Equals(CSEnumProfileWorkspace.workspace_default.ToString()))
                //    getimagelist = DataAdapter.Instance.DataCache.Profile.Option_GetValueBoolean(CSEnumOptions.queryshowimagelist);

                //bool getfulltextpreview = true;

                //bool usesearchengine = false;
                int querylistsize = int.Parse(DataAdapter.Instance.DataCache.Profile.Option_GetValue(CSEnumOptions.querylistsize));
                int queryskipcount = 0;

                // TS 10.09.14
                //bool getparents = DataAdapter.Instance.DataCache.Profile.Option_GetValueBoolean(CSEnumOptions.querygetparents);
                //bool getparents = false;
                //CSOption opt = DataAdapter.Instance.DataCache.Profile.Option_GetOption_LevelComponentsWorkspace(CSEnumOptions.querygetparents, this.Workspace);
                //if (opt != null && opt.value != null && opt.value.Length > 0)
                //    getparents = bool.Parse(opt.value);

                // TS 20.01.16 nicht mehr mit "*" suchen sondern nur noch die konkret ausgewählten indizes aus trefferliste und treeview verwenden
                //List<string> displayproperties = _Query_GetDisplayProperties(true, true);

                // queryproperties aufbereiten
                foreach (string searchrepository in repositoryquerytokens.Keys)
                {
                    List<CSQueryToken> querytokens = null;
                    if (repositoryquerytokens.TryGetValue(searchrepository, out querytokens))
                    {
                        List<CSOrderToken> orderby = null;
                        // TS 12.09.14
                        // if (!repositoryordertokens.TryGetValue(searchrepository, out orderby))
                        if (repositoryordertokens == null || !repositoryordertokens.TryGetValue(searchrepository, out orderby))
                        {
                            orderby = new List<CSOrderToken>();
                            CSOrderToken firstorderby = new CSOrderToken();
                            firstorderby.propertyname = "cmis:objectId";
                            firstorderby.orderby = CSEnumOrderBy.asc;
                            orderby.Add(firstorderby);
                        }
                        // TS 21.01.16
                        //CSQueryProperties queryproperties = _QueryCreateQueryPropertiesFromValues(searchrepository, new List<string>(), querytokens, orderby, autotruncate, queryskipcount);
                        CSQueryProperties queryproperties = _QueryCreateQueryPropertiesFromValues(searchrepository, displayproperties, querytokens, orderby, autotruncate, queryskipcount, usesearchengine);
                        querypropertieslist.Add(queryproperties);
                    }
                }

                // TS 20.03.14 suchbedingungen in datacache_objects speichern
                // TS 24.03.14 je nach schalter nur beim ersten mal
                if (!HasFixedQuery || DataAdapter.Instance.DataCache.Objects(this.Workspace).Query_LastQueryValues.Count == 0)
                    DataAdapter.Instance.DataCache.Objects(this.Workspace).Query_LastQueryValues = querypropertieslist;

                // TS 18.03.14 genereller umbau
                DataAdapter.Instance.DataProvider.Query(querypropertieslist, getparents, getfulltextpreview, getimagelist, usesearchengine, querylistsize, targetworkspace,
                                                        DataAdapter.Instance.DataCache.Rights.UserPrincipal, new CallbackAction(Query_Done, querymode, callback));
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // TS 29.01.14 noch ein weiteres callback dazu wg. query_deleted im child-window
        private void _Query_Done() { Query_Done(QueryMode.QueryOther, null); }

        private void _Query_Done(QueryMode querymode) { Query_Done(querymode, null); }

        private void _Query_Done(CallbackAction callback) { Query_Done(QueryMode.QueryOther, callback); }

        public void Query_Done(QueryMode querymode, CallbackAction callback)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                // TS 22.03.12 Vorbelegungen
                DataAdapter.Instance.DataCache.Objects(Workspace).QueryResultCount = 0;
                DataAdapter.Instance.DataCache.Objects(Workspace).QueryHasMoreResults = false;

                DataAdapter.Instance.DataCache.Objects(this.Workspace).QuerySkipStringList.Clear();
                if (DataAdapter.Instance.DataCache.ResponseStatus.success && DataAdapter.Instance.DataCache.Objects(Workspace).ObjectList_QueryResults.Count > 0)
                {
                    DataAdapter.Instance.DataCache.Objects(Workspace).RemoveEmptyQueryObjects(true);

                    // TS 21.05.15 das hier nach oben damit erst die properties gesetzt sind
                    // bevor setselectedobject gemacht wird weil sonst das paging ggf. probleme hat weil es nicht die korrekten werte kennt
                    // TS 22.03.12 in den results dasjenige suchen welches "carriesQueryInfos" enthält
                    foreach (IDocumentOrFolder obj in DataAdapter.Instance.DataCache.Objects(Workspace).ObjectList_QueryResults)
                    {
                        if (obj.QueryCarriesInfos)
                        {
                            if (obj.DZColl_QueryDocsXML.Length > 0)
                            {
                                DataAdapter.Instance.DataCache.Objects(Workspace).QueryCollectionURL = new Uri(obj.DZColl_QueryDocsURL, UriKind.RelativeOrAbsolute);
                                DataAdapter.Instance.DataCache.Objects(Workspace).QueryCollectionXML = obj.DZColl_QueryDocsXML;
                            }

                            DataAdapter.Instance.DataCache.Objects(Workspace).QueryResultCount = obj.QueryResultCount;
                            DataAdapter.Instance.DataCache.Objects(Workspace).QueryHasMoreResults = obj.QueryHasMoreResults;

                            QueryMode tmpquerymode = querymode;
                            if (querymode == QueryMode.QueryMore)
                                tmpquerymode = QueryMode.QueryNext;
                            _Query_Done_WriteSkipCount(tmpquerymode, obj);
                            // -------------------------------------------------------
                        }                      
                    }

                    // TS 30.03.15
                    if (querymode == QueryMode.QueryMore)
                    {
                        // {string[1]}
                        // [0]: "300"

                        // TS 09.03.16 blöde stelle, wenn mehrere skips vorhanden sind nimmt er trotzdem nur den ersten statt der summe
                        // letztlich führt das SetSelectedObject dazu, daß die EDP_Gesamtsicht beim Weiterlesen nicht mehr stoppt und immer wieder weiterliest
                        // also raus damit, zumal mit aktuellem paging es sowieso anders gelöst werden sollte
                        //int lastquerycount = 0;
                        //string[] slastcount = _Query_GetSkipCount(QueryMode.QueryNext);
                        //if (slastcount != null && slastcount.Length > 0)
                        //    lastquerycount = int.Parse(slastcount[0]);
                        //if (DataAdapter.Instance.DataCache.Objects(Workspace).ObjectList_QueryResults.Count > lastquerycount)
                        //    this.SetSelectedObject(DataAdapter.Instance.DataCache.Objects(Workspace).ObjectList_QueryResults[lastquerycount].objectId);
                        querymode = QueryMode.QueryNext;
                    }
                    else if (Workspace != CSEnumProfileWorkspace.workspace_searchoverall)
                    {
                        this.SetSelectedObject(DataAdapter.Instance.DataCache.Objects(Workspace).ObjectList_QueryResults[0].objectId);
                    }
                    // TS 30.11.16 im searchoverall workspace auch den ersten treffer anzeigen wenn er im aktuell gewählten archiv ist
                    else if (Workspace == CSEnumProfileWorkspace.workspace_searchoverall)
                    {
                        IDocumentOrFolder dummy = DataAdapter.Instance.DataCache.Objects(Workspace).ObjectList_QueryResults[0];
                        if (dummy.RepositoryId.Equals(DataAdapter.Instance.DataCache.Rights.ChoosenRepository.repositoryId))
                        {
                            CSEnumProfileWorkspace defws = CSEnumProfileWorkspace.workspace_default;
                            if (!DataAdapter.Instance.DataCache.Objects(defws).ExistsObject(dummy.objectId))
                            {
                                //CallbackAction callBack = new CallbackAction(_QuerySearchOverAll_CopySiblingsToDefault, dummy.objectId);
                                _QuerySearchOverAll_CopySiblingsToDefault(dummy.objectId);                                
                                //DataAdapter.Instance.Processing(defws).QuerySingleObjectById(dummy.RepositoryId, dummy.objectId, false, true, defws, callBack);
                            }
                            else
                            {
                                DataAdapter.Instance.Processing(defws).SetSelectedObject(dummy.objectId);
                            }
                        }
                    }
                }

                // TS 26.09.14 wenn nix gefunden dann ein notify senden damit die buttons sich wieder sauber schalten
                // TS 14.10.15 erweitertes notify damit sich auch die pager in den listen sauber setzen
                //if (dataadapter.instance.datacache.objects(workspace).queryresultcount == 0)
                //{
                //    dataadapter.instance.informobservers("cantoggle_query");
                //}
                DataAdapter.Instance.InformObservers();

                if (callback != null) callback.Invoke();
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        private void _QuerySearchOverAll_CopySiblingsToDefault(string objectID)
        {
            List<IDocumentOrFolder> addList = new List<IDocumentOrFolder>();
            IDocumentOrFolder selObj = DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_searchoverall).GetObjectById(objectID);

            // Collect all siblings of the selected item
            foreach (IDocumentOrFolder itemInCache in DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_searchoverall).ObjectList)
            {
                if (itemInCache.parentId.Equals(selObj.parentId) 
                    && !DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_default).ExistsObject(itemInCache.objectId) 
                    && !itemInCache.objectId.Equals(selObj.objectId)
                    && itemInCache.RepositoryId.Equals(selObj.RepositoryId) 
                    && selObj.structLevel < Statics.Constants.STRUCTLEVEL_07_AKTE)
                {
                    addList.Add(itemInCache);
                }
            }

            // Insert CopyToWorkspace Method here
            CallbackAction cb_copyObject = new CallbackAction(SendCopyWorkspace, addList, CSEnumProfileWorkspace.workspace_default, objectID);
            CallbackAction cb_SetSelectedObject = new CallbackAction(DataAdapter.Instance.Processing(CSEnumProfileWorkspace.workspace_default).SetSelectedObject, objectID, 0, cb_copyObject);

            // Query for the object, to get it's parents
            if (!DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_default).ExistsObject(selObj.objectId))
            {
                DataAdapter.Instance.Processing(CSEnumProfileWorkspace.workspace_default).QuerySingleObjectById(selObj.RepositoryId, selObj.objectId, false, true, CSEnumProfileWorkspace.workspace_default, cb_SetSelectedObject);
            }
            else
            {
                if (!DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_default).Object_Selected.objectId.Equals(selObj.objectId))
                {
                    DataAdapter.Instance.Processing(CSEnumProfileWorkspace.workspace_default).SetSelectedObject(selObj.objectId, 0, cb_copyObject);
                }
            }
        }

        private void SendCopyWorkspace(object objectsToChange, object targetWorkspace, object origSelectedID)
        {
            string origID = (string)origSelectedID;
            List<IDocumentOrFolder> cast_objectsToChange = (List<IDocumentOrFolder>)objectsToChange;
            CSEnumProfileWorkspace cast_targetWorkspace = (CSEnumProfileWorkspace)targetWorkspace;
            CallbackAction callBack = new CallbackAction(SetSelectedObject, origID);

            // Apply Fulltext-Values to the origin selected object and select the first matching page
            IDocumentOrFolder newObject = DataAdapter.Instance.DataCache.Objects(cast_targetWorkspace).GetObjectById(origID);
            IDocumentOrFolder oldObject = DataAdapter.Instance.DataCache.Objects(CSEnumProfileWorkspace.workspace_searchoverall).GetObjectById(origID);
            newObject.FulltextProperty = oldObject.FulltextProperty;

            DataAdapter.Instance.DataProvider.ChangeWorkspace(cast_objectsToChange, cast_targetWorkspace, DataAdapter.Instance.DataCache.Rights.UserPrincipal, callBack);
        }

        private void _QueryDepending(string in_folder, List<CSOrderToken> givenorderby, CSEnumProfileWorkspace targetworkspace)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                // vorbelegungen
                DataAdapter.Instance.DataCache.Objects(Workspace).QueryCollectionURL = new Uri("", UriKind.RelativeOrAbsolute);
                DataAdapter.Instance.DataCache.Objects(Workspace).QueryCollectionXML = "";
                bool autotruncate = false;
                bool getimagelist = false;
                //bool getfulltextpreview = true;
                //bool getfulltextpreview = false;
                bool getfulltextpreview = true;

                bool usesearchengine = false;
                int querylistsize = int.Parse(DataAdapter.Instance.DataCache.Profile.Option_GetValue(CSEnumOptions.querytreesize));
                int queryskipcount = 0;
                bool getparents = false;

                // baut querytokens aus den abgeholten werten
                List<CServer.CSQueryToken> querytokens = new List<CServer.CSQueryToken>();

                if (in_folder != null && in_folder.Length > 0)
                {
                    CSQueryToken token = new CSQueryToken();

                    // TS 03.01.18 hier: änderung auf IN_TREE
                    //token.propertyname = "IN_FOLDER";
                    // insgesamt 4 Stellen im SL Code
                    // 1. Query_EDP_Gesamt: Suche in Posteingang => bleibt IN_FOLDER
                    // 2. Query_EDP_Gesamt: Suche in Maileingang => bleibt IN_FOLDER
                    // 3. Query_Depending (hier):                => wird   IN_TREE
                    // 4. _Query_Proceed:                        => wird   IN_TREE
                    token.propertyname = "IN_TREE";

                    token.propertyvalue = in_folder;
                    token.propertytype = enumPropertyType.@string;
                    querytokens.Add(token);
                }

                // sortierung
                List<CServer.CSOrderToken> orderby = new List<CSOrderToken>();
                if (givenorderby != null && givenorderby.Count > 0)
                {
                    foreach (CSOrderToken token in givenorderby)
                        orderby.Add(token);
                }
                else
                {
                    CSOrderToken firstorderby = new CSOrderToken();
                    firstorderby.propertyname = "cmis:objectId";
                    firstorderby.orderby = CSEnumOrderBy.asc;
                    orderby.Add(firstorderby);
                }

                // TS 18.03.14 genereller umbau
                string searchrepository = DataAdapter.Instance.DataCache.Repository(Workspace).RepositoryInfo.repositoryId;
                List<CSQueryProperties> querypropertieslist = new List<CSQueryProperties>();

                // TS 20.01.16 nicht mehr mit "*" suchen sondern nur noch die konkret ausgewählten indizes aus trefferliste und treeview verwenden
                List<string> displayproperties = _Query_GetDisplayProperties(false, true);

                // TS 21.01.16
                //CSQueryProperties queryproperties = _QueryCreateQueryPropertiesFromValues(searchrepository, new List<string>(), querytokens, orderby, autotruncate, queryskipcount);
                CSQueryProperties queryproperties = _QueryCreateQueryPropertiesFromValues(searchrepository, displayproperties, querytokens, orderby, autotruncate, queryskipcount, false);
                querypropertieslist.Add(queryproperties);

                // TS 18.03.14 genereller umbau
                DataAdapter.Instance.DataProvider.Query(querypropertieslist, getparents, getfulltextpreview, getimagelist, usesearchengine, querylistsize, targetworkspace,
                                                        DataAdapter.Instance.DataCache.Rights.UserPrincipal);
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        private void _QueryStatics(List<cmisTypeContainer> staticlevels, Dictionary<string, string> staticqueries, List<CSOrderToken> staticorderby, CSEnumProfileWorkspace targetworkspace)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                // vorbelegungen
                //DataAdapter.Instance.DataCache.Objects(Workspace).QueryCollectionURL = new Uri("", UriKind.RelativeOrAbsolute);
                //DataAdapter.Instance.DataCache.Objects(Workspace).QueryCollectionXML = "";
                bool autotruncate = false;
                bool getimagelist = false;
                bool getfulltextpreview = false;
                bool usesearchengine = false;
                int querylistsize = 999;
                int queryskipcount = 0;
                bool getparents = false;

                List<CServer.CSQueryToken> querytokens = new List<CServer.CSQueryToken>();

                foreach (cmisTypeContainer container in staticlevels)
                {
                    CSQueryToken token = new CSQueryToken();
                    token.propertyname = "";
                    token.propertyvalue = "";
                    token.propertyreptypeid = container.type.id;
                    querytokens.Add(token);
                }

                // TS 14.03.14 zusätzliche statische suchbedingungen
                foreach (string key in staticqueries.Keys)
                {
                    CSQueryToken token2 = new CSQueryToken();
                    token2.propertyname = key;
                    token2.propertyvalue = staticqueries[key];

                    // TS 15.04.15
                    //token2.propertyvalue = _QueryProcessValueReplacements(token2.propertyvalue);
                    token2.propertyvalue = ValueReplacements.ProcessValueReplacements(token2.propertyvalue, DataAdapter.Instance.DataCache.Rights.UserPrincipal);

                    token2.propertyvalue = token2.propertyvalue.Replace("|", Constants.QUERY_OPERATOR_OR);
                    cmisTypeContainer tc = DataAdapter.Instance.DataCache.Repository(this.Workspace).GetTypeContainerForPropertyName(key);
                    if (tc != null)
                    {
                        token2.propertyreptypeid = tc.type.id;
                        querytokens.Add(token2);
                    }
                }

                // sortierung
                List<CServer.CSOrderToken> orderby = new List<CSOrderToken>();
                if (staticorderby != null && staticorderby.Count > 0)
                {
                    foreach (CSOrderToken token in staticorderby)
                        orderby.Add(token);
                }
                else
                {
                    CSOrderToken firstorderby = new CSOrderToken();
                    firstorderby.propertyname = "cmis:objectId";
                    firstorderby.orderby = CSEnumOrderBy.asc;
                    orderby.Add(firstorderby);
                }

                // TS 18.03.14 genereller umbau
                string searchrepository = DataAdapter.Instance.DataCache.Repository(Workspace).RepositoryInfo.repositoryId;
                List<CSQueryProperties> querypropertieslist = new List<CSQueryProperties>();

                // TS 20.01.16 nicht mehr mit "*" suchen sondern nur noch die konkret ausgewählten indizes aus trefferliste und treeview verwenden
                List<string> displayproperties = _Query_GetDisplayProperties(true, false);

                // TS 21.01.16
                //CSQueryProperties queryproperties = _QueryCreateQueryPropertiesFromValues(searchrepository, new List<string>(), querytokens, orderby, autotruncate, queryskipcount);
                CSQueryProperties queryproperties = _QueryCreateQueryPropertiesFromValues(searchrepository, displayproperties, querytokens, orderby, autotruncate, queryskipcount, false);
                querypropertieslist.Add(queryproperties);

                // TS 18.03.14 genereller umbau
                DataAdapter.Instance.DataProvider.Query(querypropertieslist, getparents, getfulltextpreview, getimagelist, usesearchengine, querylistsize, targetworkspace,
                                                        DataAdapter.Instance.DataCache.Rights.UserPrincipal, new CallbackAction(_QueryStatics_Done));
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        private void _QueryStatics_Done()
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                if (DataAdapter.Instance.DataCache.ResponseStatus.success)
                {
                    DataAdapter.Instance.DataCache.Objects(Workspace).RemoveEmptyQueryObjects(true);
                    if (DataAdapter.Instance.DataCache.Objects(Workspace).ObjectList.Count > 1)
                    {
                        // TS 14.03.14
                        if (DefaultStaticFolderId.Length > 0 && DataAdapter.Instance.DataCache.Objects(Workspace).ObjectIdList_Statics.Contains(DefaultStaticFolderId))
                        {
                            IDocumentOrFolder defaultfolder = DataAdapter.Instance.DataCache.Objects(Workspace).GetObjectById(DefaultStaticFolderId);
                            CallbackAction callback = new CallbackAction(SetSelectedObject, DefaultStaticFolderId);
                            this.GetObjectsUpDown(defaultfolder, true, callback);
                        }
                        else
                            this.SetSelectedObject(DataAdapter.Instance.DataCache.Objects(Workspace).ObjectList[1].objectId);
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        /// <summary>
        /// querymode 0 = neue suche / 1 = nächste treffer / 2 = vorherige treffer
        /// </summary>
        /// <param name="querymode"></param>
        /// <param name="in_folder"></param>
        /// <param name="targetworkspace"></param>
        private void _QueryDeleted(QueryMode querymode, CSEnumProfileWorkspace targetworkspace, CallbackAction callback_done)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                int querylistsize = int.Parse(DataAdapter.Instance.DataCache.Profile.Option_GetValue(CSEnumOptions.querylistsize));
                int skipcount = 0;
                // TS 01.07.14 skipcount nur setzen wenn nicht new
                if (querymode != QueryMode.QueryNew)
                {
                    string[] tokens = _Query_GetSkipCount(querymode);
                    if (tokens.Count() > 0)
                        skipcount = Int32.Parse(tokens[0]);
                }

                // alle repositories abfragen: "0_mandantid"
                string searchrepository = "0_" + DataAdapter.Instance.DataCache.Rights.ChoosenMandant.id;
                //DataAdapter.Instance.DataProvider.QueryDeleted(searchrepository, querylistsize, skipcount, targetworkspace, DataAdapter.Instance.DataCache.Rights.UserPrincipal,
                //    new CallbackAction(_Query_Done));
                CallbackAction callback = new CallbackAction(Query_Done, querymode, new CallbackAction(_QueryDeleted_Done, "", callback_done));
                DataAdapter.Instance.DataProvider.QueryDeleted(searchrepository, querylistsize, skipcount, targetworkspace, DataAdapter.Instance.DataCache.Rights.UserPrincipal, callback);
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        private void _QueryDeleted_Done(string tmp, CallbackAction callback_done)
        {
            try
            {
                DataAdapter.Instance.InformObservers(this.Workspace);
                if (callback_done != null)
                {
                    callback_done.Invoke();
                }
            }
            catch (Exception ex) { if (ex.Message.Length > 0) { Log.Log.Debug(ex.Message); } }
        }


        private void _QueryLastEdited(QueryMode querymode, CSEnumProfileWorkspace targetworkspace, CallbackAction callback_done)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                int querylistsize = int.Parse(DataAdapter.Instance.DataCache.Profile.Option_GetValue(CSEnumOptions.querylistsize));
                int skipcount = 0;
                // TS 01.07.14 skipcount nur setzen wenn nicht new
                if (querymode != QueryMode.QueryNew)
                {
                    string[] tokens = _Query_GetSkipCount(querymode);
                    if (tokens.Count() > 0)
                        skipcount = Int32.Parse(tokens[0]);
                }

                // alle repositories abfragen: "0_mandantid"
                string mandantID = DataAdapter.Instance.DataCache.Rights.ChoosenMandant.id;
                string searchrepository = "[0_" + mandantID + "," + "-1_" + mandantID + "," + "-2_" + mandantID + "," + "-5_" + mandantID + "," + "-8_" + mandantID + "]";
                //DataAdapter.Instance.DataProvider.QueryDeleted(searchrepository, querylistsize, skipcount, targetworkspace, DataAdapter.Instance.DataCache.Rights.UserPrincipal,
                //    new CallbackAction(_Query_Done));
                CallbackAction callback = new CallbackAction(Query_Done, querymode, new CallbackAction(_QueryLastEdited_Done, "", callback_done));
                DataAdapter.Instance.DataProvider.QueryLastEdited(searchrepository, querylistsize, skipcount, targetworkspace, DataAdapter.Instance.DataCache.Rights.UserPrincipal, callback);
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        private void _QueryLastEdited_Done(string tmp, CallbackAction callback_done)
        {
            try
            {
                DataAdapter.Instance.InformObservers(this.Workspace);
                if (callback_done != null)
                {
                    callback_done.Invoke();
                }
            }
            catch (Exception ex) { if (ex.Message.Length > 0) { Log.Log.Debug(ex.Message); } }
        }


        // =============================================================================

        public void RefreshQueryCount()
        {
            try
            {
                char splitter = (";".ToCharArray())[0];
                char splitter2nd = ("=".ToCharArray())[0];
                string queryInfo = DataAdapter.Instance.DataCache.Info[CSEnumInformationId.QueryCountData];

                DataAdapter.Instance.DataCache.Info.QueryIndexCount().Clear();
                if (queryInfo != null)
                {
                    string[] queryInformation = queryInfo.Split(splitter);
                    foreach(string info in queryInformation)
                    {
                        if (info.Contains("="))
                        {
                            string infoName = info.Split(splitter2nd)[0];
                            string infoValue = info.Split(splitter2nd)[1];
                            int structLvl = infoName.Contains("INDEX_") ? int.Parse(infoName.Substring(6)) : 99;
                            DataAdapter.Instance.DataCache.Info.QueryIndexCount().Add(structLvl, infoValue);
                        }
                    }
                }
            }
            catch (System.Exception ex) { if (ex.Message.Length > 0) { Log.Log.Debug(ex.Message); } }
        }

        // =============================================================================

        /// <summary>
        /// holt im AKTUELLEN workspace definierte datenfilter
        /// </summary>
        private void _Query_AddDataFilter(ref List<string> searchrepositories, ref Dictionary<string, string> queryvalues)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                CSProfileWorkspace ws = DataAdapter.Instance.DataCache.Profile.Profile_GetProfileWorkspaceFromType(this.Workspace);
                if (ws != null && ws.datafilters != null)
                {
                    foreach (CSProfileDataFilter df in ws.datafilters)
                    {
                        if (df.type == CSEnumProfileDataFilterType.ws_query)
                        {
                            CSEnumProfileDataFilterMode mode = df.mode;
                            string indexid = df.indexid.ToUpper();
                            string selectedvalues = df.selectedvalues;
                            if (selectedvalues.Length > 0)
                            {
                                char splitter = ("|".ToCharArray())[0];
                                string[] tokens = selectedvalues.Split(splitter);
                                string values = "";
                                for (int i = 0; i < tokens.Length; i++)
                                {
                                    if (indexid.ToUpper().Equals("CMIS:REPOSITORY"))
                                    {
                                        string reptry = tokens[i] + "_" + DataAdapter.Instance.DataCache.Rights.UserPrincipal.mandantid;
                                        if (!searchrepositories.Contains(reptry))
                                            searchrepositories.Add(reptry);
                                    }
                                    else if (!indexid.ToUpper().StartsWith("CMIS:"))
                                    {
                                        if (values.Length > 0)
                                            values = values + Constants.QUERY_OPERATOR_OR;
                                        values = values + tokens[i];
                                    }
                                }
                                if (values.Length > 0)
                                    queryvalues.Add(indexid, values);
                            }
                        }
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        private List<string> _Query_GetDisplayProperties(bool gettreeproperties, bool getlistproperties) { return _Query_GetDisplayProperties(gettreeproperties, getlistproperties, this); }

        /// <summary>
        /// holt ausgewählte spalten von Tree aus AKTUELLEM workspace und List ggf. aus overall
        /// </summary>
        private List<string> _Query_GetDisplayProperties(bool gettreeproperties, bool getlistproperties, IProcessing overallprocessing)
        {
            List<cmisTypeContainer> types = DataAdapter.Instance.DataCache.Repository(this.Workspace).TypeDescendants;

            List<string> queryresultcolumns = new List<string>();
            if (gettreeproperties && TreeDisplayProperties.Count > 0)
            {
                foreach (string property in TreeDisplayProperties)
                {
                    // TS 25.01.16 nur die typen dieses workspace verwenden
                    cmisTypeContainer container = DataAdapter.Instance.DataCache.Repository(this.Workspace).GetTypeContainerForPropertyName(property);
                    if (container == null) container = DataAdapter.Instance.DataCache.Repository_FindTypeContainerForPropertyName(property);
                    if (container != null && !queryresultcolumns.Contains(property))
                    {
                        foreach (cmisTypeContainer type in types)
                        {
                            if (container.type.id.Equals(type.type.id))
                            {
                                queryresultcolumns.Add(property);
                                break;
                            }
                        }
                    }
                }
            }
            if (getlistproperties && overallprocessing.ListDisplayProperties.Count > 0)
            {
                foreach (string property in overallprocessing.ListDisplayProperties)
                {
                    // nur den jeweils abgefragten typ übergeben
                    cmisTypeContainer container = DataAdapter.Instance.DataCache.Repository(this.Workspace).GetTypeContainerForPropertyName(property);
                    if (container == null) container = DataAdapter.Instance.DataCache.Repository_FindTypeContainerForPropertyName(property);
                    if (container != null && !queryresultcolumns.Contains(property))
                    {
                        queryresultcolumns.Add(property);
                    }
                }
            }
            return queryresultcolumns;
        }


        private void _Query_GetSortDataFilter(ref List<CSOrderToken> queryorderby, bool isquery) { _Query_GetSortDataFilter(ref queryorderby, isquery, this); }

        /// <summary>
        /// liefert die CSOrderTokens aus dem Workspace + je nach Parameter(isquery) die Query- oder Struct- Order Tokens dazu
        /// workspace und tree sortierungen werden aus aktuellem geholt, liste ggf. aus overall
        /// </summary>
        private void _Query_GetSortDataFilter(ref List<CSOrderToken> queryorderby, bool isquery, IProcessing overallprocessing)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                // sortierungen aus workspace haben vorrang
                if (WorkspaceOrderTokens.Count > 0)
                {
                    foreach (CSOrderToken orderby in WorkspaceOrderTokens)
                    {
                        // TS 28.06.13 das muss hier dazu sonst nimmt er immer den default der klasse = asc
                        orderby.orderbySpecified = true;
                        queryorderby.Add(orderby);
                    }
                }
                // TS 01.10.13 die sortierungen aus processing (gesetzt von listview) wenn es eine "echte" Suche ist
                // TS 14.11.16
                // else if (isquery && ListOrderTokens.Count > 0)
                else if (isquery && overallprocessing.ListOrderTokens.Count > 0)
                {
                    foreach (CSOrderToken orderby in overallprocessing.ListOrderTokens)
                    {
                        // TS 28.06.13 das muss hier dazu sonst nimmt er immer den default der klasse = asc
                        orderby.orderbySpecified = true;
                        queryorderby.Add(orderby);
                    }
                }
                // die sortierung aus processing (gesetzt von treeview) wenn es KEINE Suche ist sondern z.B. GetChildren
                else if (!isquery && TreeOrderTokens.Count > 0)
                {
                    foreach (CSOrderToken orderby in TreeOrderTokens)
                    {
                        // TS 28.06.13 das muss hier dazu sonst nimmt er immer den default der klasse = asc
                        orderby.orderbySpecified = true;
                        queryorderby.Add(orderby);
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        // TS 26.11.12 unterscheidung fulltext oder nicht
        private List<CServer.CSQueryToken> _Query_CreateQueryTokens(Dictionary<string, string> queryvalues, bool fulltextvalues)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            List<CServer.CSQueryToken> querytokens = new List<CSQueryToken>();
            string realpropname = "";
            try
            {
                // TS 01.04.14
                bool levelforced = false;
                // TS 04.07.14
                List<string> queryvaluestoremove = new List<string>();
                Dictionary<string, string> queryvaluestoadd = new Dictionary<string, string>();
                foreach (string propname in queryvalues.Keys)
                {
                    // TS 26.11.12
                    realpropname = propname;
                    // TS 26.11.12
                    if ((!fulltextvalues && !realpropname.StartsWith(Constants.FULLTEXT_PREFIX)) || (fulltextvalues && realpropname.StartsWith(Constants.FULLTEXT_PREFIX)))
                    {
                        CSQueryToken token = new CSQueryToken();
                        // TS 26.11.12
                        if (realpropname.StartsWith(Constants.FULLTEXT_PREFIX))
                            realpropname = realpropname.Substring(Constants.FULLTEXT_PREFIX.Length);
                        // TS 04.07.14
                        if (realpropname.EndsWith(Constants.QUERY_RANGEATTRIBFROM))
                            realpropname = realpropname.Replace(Constants.QUERY_RANGEATTRIBFROM, "");
                        if (realpropname.EndsWith(Constants.QUERY_RANGEATTRIBTO))
                            realpropname = realpropname.Replace(Constants.QUERY_RANGEATTRIBTO, "");
                        token.propertyname = realpropname;

                        // If the value depends on a datasource, force to not autotruncate
                        string listIndexID = DataAdapter.Instance.DataCache.Meta.GetListIdFromIndexId(realpropname);
                        if (listIndexID != null && listIndexID.Length > 0)
                        {
                            token.ignoreautotruncate = true;
                        }

                        // TS 02.07.13 parentfolders weglassen da sonderbehandlung wg. mailmill
                        if (!realpropname.ToUpper().Equals("PARENTFOLDERS"))
                        {
                            string propvalue = "";
                            queryvalues.TryGetValue(propname, out propvalue);

                            // TS 23.06.14 prüfen ob eine range vorliegt
                            if (propname.Contains(Constants.QUERY_RANGEATTRIBFROM))
                            {
                                // wenn eine rangefrom-property gefunden wurde
                                if (propvalue != null && propvalue.Length > 0)
                                    propvalue = ">=" + propvalue;
                                // wenn auch eine rangeto existiert dann beide hier wegschreiben
                                if (queryvalues.Keys.Contains(propname.Replace(Constants.QUERY_RANGEATTRIBFROM, Constants.QUERY_RANGEATTRIBTO)))
                                {
                                    string proptovalue = "";
                                    queryvalues.TryGetValue(propname.Replace(Constants.QUERY_RANGEATTRIBFROM, Constants.QUERY_RANGEATTRIBTO), out proptovalue);
                                    if (proptovalue != null && proptovalue.Length > 0)
                                        propvalue = propvalue + Constants.QUERY_OPERATOR_AND + " <=" + proptovalue;
                                }
                            }
                            else if (propname.Contains(Constants.QUERY_RANGEATTRIBTO))
                            {
                                // nur wenn keine rangefrom existiert dann hier wegschreiben (sonst wurde es bereits dort abgeholt)
                                if (!queryvalues.Keys.Contains(propname.Replace(Constants.QUERY_RANGEATTRIBTO, Constants.QUERY_RANGEATTRIBFROM)))
                                {
                                    if (propvalue != null && propvalue.Length > 0)
                                        propvalue = "<=" + propvalue;
                                }
                                else
                                    propvalue = "";
                            }
                            else
                            {
                                // wenn bereits über die range abgehandelt wurde
                                if (queryvalues.Keys.Contains(propname + Constants.QUERY_RANGEATTRIBFROM) || queryvalues.Keys.Contains(propname + Constants.QUERY_RANGEATTRIBTO))
                                    propvalue = "";
                            }

                            // TS 22.01.14 Ersetzungen
                            // TS 15.04.15
                            //propvalue = _QueryProcessValueReplacements(propvalue);
                            propvalue = ValueReplacements.ProcessValueReplacements(propvalue, DataAdapter.Instance.DataCache.Rights.UserPrincipal);

                            token.propertyvalue = propvalue;
                            cmisTypeContainer container = DataAdapter.Instance.DataCache.Repository(Workspace).GetTypeContainerForPropertyName(realpropname);
                            // TS 28.02.14
                            //if (container == null) container = DataAdapter.Instance.DataCache.Repository_FindTypeContainerForPropertyName(realpropname);

                            // TS 01.04.14 sonderbehandlung für abgefragte cmis_properties:
                            // wenn nur eine property abgefragt wird und diese ist eine cmis property dann immer nach logischem dokument suchen !!!
                            //if (propname.StartsWith(Constants.CMISPROPERTY) && queryvalues.Count == 1)
                            if (realpropname.Equals(CSEnumCmisProperties.cmis_lastModifiedBy)
                                ||
                                realpropname.Equals(CSEnumCmisProperties.cmis_lastModifiedBy.ToString().Replace("_", ":")))
                            {
                                container = DataAdapter.Instance.DataCache.Repository(this.Workspace).GetTypeContainerForStructLevel(Constants.STRUCTLEVEL_09_DOKLOG);
                                levelforced = true;
                            }

                            // TS 19.03.14
                            if (container == null)
                            {
                                container = DataAdapter.Instance.DataCache.Repository(this.Workspace).GetTypeContainerForStructLevel(Constants.STRUCTLEVEL_11_FULLTEXT);
                            }

                            if (container != null)
                            {
                                token.propertytype = DataAdapter.Instance.DataCache.Repository(Workspace).GetPropertyTypeForPropertyName(realpropname, container);
                                // TS 28.04.14 das hat gefehlt
                                token.propertytypeSpecified = true;

                                // TS 14.03.12 tablename nur dranhängen wenn nicht direkte cmis:property (z.B. cmis:objectId) abgefragt wird
                                // WICHTIG: damit überhaupt solche festen columns abgefragt werden können muss
                                // 1. das binding auf 2 stehen
                                // 2. das binding nicht an eine der DocumentOrFolder.standardProperty (z.B. objectId) gebunden sein, da diese nur getter und keine setter haben
                                // und somit in beiden fällen gar keine suchbedingung aus der maske abgeholt wird

                                // TS 26.07.12 testweise rausgenommen für archivübergreifende suche
                                //if (!realpropname.Contains("cmis:"))
                                //if (!realpropname.Contains(Constants.CMISPROPERTY + ":"))
                                if (!realpropname.Contains(Constants.CMISPROPERTY + ":") || levelforced)
                                    token.propertyreptypeid = container.type.id;

                                // TS 04.07.14 nur wenn eetwas drinsteht
                                if (token.propertyvalue != null && token.propertyvalue.Length > 0)
                                    querytokens.Add(token);
                            }
                        }
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
            return querytokens;
        }

        /// <summary>
        /// baut ein queryproperties objekt aus den angegebenen werten
        /// </summary>
        private CSQueryProperties _QueryCreateQueryPropertiesFromValues(string searchrepository, List<string> searchcolumns,
                                                                        List<CServer.CSQueryToken> searchconditions, List<CServer.CSOrderToken> orderby,
                                                                        bool autotruncate, long skipcount, bool isFulltext)
        {
            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            CSQueryProperties queryproperties = new CSQueryProperties();
            bool isPhonetic = DataAdapter.Instance.DataCache.Profile.Option_GetValueBoolean(CSEnumOptions.phoneticsearch);
            try
            {
                // SearchColumns to Array
                string[] asearchcolumns = new string[searchcolumns.Count];
                for (int i = 0; i < searchcolumns.Count; i++) { asearchcolumns.SetValue(searchcolumns[i], i); }

                // QueryTokens to Array
                CServer.CSQueryToken[] asearchconditions = new CServer.CSQueryToken[searchconditions.Count];
                for (int i = 0; i < searchconditions.Count; i++)
                {
                    asearchconditions.SetValue(searchconditions[i], i);

                    // Add Phonetic-Behaviour to the Search-Condition
                    if(isPhonetic && isFulltext && !searchconditions[i].propertyvalue.StartsWith("KP("))
                    {
                        searchconditions[i].propertyvalue = "KP(" + searchconditions[i].propertyvalue + ")";
                    }
                }

                // Order-Tokens to Array
                CServer.CSOrderToken[] aorderby = new CServer.CSOrderToken[orderby.Count];
                for (int i = 0; i < orderby.Count; i++) { aorderby.SetValue(orderby[i], i); }

                // Build the QueryProeprtie
                queryproperties.searchcolumns = asearchcolumns;
                queryproperties.searchconditions = asearchconditions;
                queryproperties.searchrepository = searchrepository;
                queryproperties.orderby = aorderby;
                queryproperties.skipcount = skipcount.ToString();
                queryproperties.autotruncate = autotruncate;
                queryproperties.autotruncateSpecified = true;
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
            return queryproperties;
        }

        /// <summary>
        /// holt skipcount aus dem AKTUELLEN workspace
        /// </summary>
        private void _Query_GetSkipCount(QueryMode querymode, List<CSQueryProperties> querypropertieslist)
        {
            //            curr	   anz	      next		coll		prev

            //00 - 20		00-00-00 | 10-03-07 | 10-03-07		10-03-07	n.a. [00-00-00]

            //20 - 40		10-03-07 | 05-05-10 | 15-08-17		15-08-17	00-00-00

            //40 - 60		15-08-17 | 00-12-08 | 15-20-25		15-20-25	10-03-07

            //60 - 80		15-20-25 | 13-02-05 | 28-22-30		28-22-30

            //        => next = letzter eintrag (collection.count - 1)
            //        => curr = vorletz eintrag (collection.count - 2)
            //        => prev = vorvorl eintrag (collection.count - 3)

            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();
            try
            {
                string[] tokens = _Query_GetSkipCount(querymode);

                for (int i = 0; i < querypropertieslist.Count; i++)
                {
                    CSQueryProperties prop = querypropertieslist[i];
                    string token = "0";
                    if (tokens.Count() > i)
                        token = tokens[i];
                    prop.skipcount = token;
                    i++;
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
        }

        private string[] _Query_GetSkipCount(QueryMode querymode)
        {
            //            curr	   anz	      next		coll		prev

            //00 - 20		00-00-00 | 10-03-07 | 10-03-07		10-03-07	n.a. [00-00-00]

            //20 - 40		10-03-07 | 05-05-10 | 15-08-17		15-08-17	00-00-00

            //40 - 60		15-08-17 | 00-12-08 | 15-20-25		15-20-25	10-03-07

            //60 - 80		15-20-25 | 13-02-05 | 28-22-30		28-22-30

            //        => next = letzter eintrag (collection.count - 1)
            //        => curr = vorletz eintrag (collection.count - 2)
            //        => prev = vorvorl eintrag (collection.count - 3)

            if (Log.Log.IsDebugEnabled) Log.Log.MethodEnter();

            string[] rettokens = new string[1];
            rettokens[0] = "0";
            try
            {
                int offset = 1;
                switch (querymode)
                {
                    case QueryMode.QueryPrev:
                        offset = 3;
                        break;

                    case QueryMode.QueryRefresh:
                        offset = 2;
                        break;

                    case QueryMode.QueryNext:
                        offset = 1;
                        break;
                }
                if (DataAdapter.Instance.DataCache.Objects(this.Workspace).QuerySkipStringList.Count > offset - 1)
                {
                    string queryskiplist = DataAdapter.Instance.DataCache.Objects(this.Workspace).QuerySkipStringList[DataAdapter.Instance.DataCache.Objects(this.Workspace).QuerySkipStringList.Count - offset];
                    if (queryskiplist != null && queryskiplist.Length > 0)
                    {
                        // TS 21.03.14
                        //if (queryskiplist.Contains("|"))
                        //{
                        //    char splitter = ("|".ToCharArray())[0];
                        //    rettokens = queryskiplist.Split(splitter);
                        //}
                        //else
                        //{
                        //    rettokens[0] = queryskiplist;
                        //}
                        rettokens = _Query_GetSkipTokens(queryskiplist);
                    }
                }
            }
            catch (Exception e) { Log.Log.Error(e); }
            if (Log.Log.IsDebugEnabled) Log.Log.MethodLeave();
            return rettokens;
        }

        /// <summary>
        /// schreibt skipcount in den AKTUELLEN workspace
        /// </summary>
        private void _Query_Done_WriteSkipCount(QueryMode querymode, IDocumentOrFolder queryinfoobject)
        {
            string newskipcount = queryinfoobject[CSEnumCmisProperties.queryResultSkipString];
            if (newskipcount != null && newskipcount.Length > 0)
            {
                switch (querymode)
                {
                    case QueryMode.QueryNew:
                        DataAdapter.Instance.DataCache.Objects(this.Workspace).QuerySkipStringList.Clear();
                        DataAdapter.Instance.DataCache.Objects(this.Workspace).QuerySkipStringList.Add(newskipcount);
                        break;

                    case QueryMode.QueryFulltextAll:
                        DataAdapter.Instance.DataCache.Objects(this.Workspace).QuerySkipStringList.Clear();
                        DataAdapter.Instance.DataCache.Objects(this.Workspace).QuerySkipStringList.Add(newskipcount);
                        break;

                    case QueryMode.QueryFulltextDetail:
                        DataAdapter.Instance.DataCache.Objects(this.Workspace).QuerySkipStringList.Clear();
                        DataAdapter.Instance.DataCache.Objects(this.Workspace).QuerySkipStringList.Add(newskipcount);
                        break;

                    case QueryMode.QueryRefresh:
                        DataAdapter.Instance.DataCache.Objects(this.Workspace).QuerySkipStringList.Clear();
                        DataAdapter.Instance.DataCache.Objects(this.Workspace).QuerySkipStringList.Add(newskipcount);
                        break;

                    case QueryMode.QuerySearchEngine:
                        DataAdapter.Instance.DataCache.Objects(this.Workspace).QuerySkipStringList.Clear();
                        DataAdapter.Instance.DataCache.Objects(this.Workspace).QuerySkipStringList.Add(newskipcount);
                        break;

                    case QueryMode.QueryPrev:
                        if (DataAdapter.Instance.DataCache.Objects(this.Workspace).QuerySkipStringList.Count > 0)
                            DataAdapter.Instance.DataCache.Objects(this.Workspace).QuerySkipStringList.RemoveAt(DataAdapter.Instance.DataCache.Objects(this.Workspace).QuerySkipStringList.Count - 1);
                        break;

                    case QueryMode.QueryNext:
                        if (DataAdapter.Instance.DataCache.Objects(this.Workspace).QuerySkipStringList.Count > 0)
                        {
                            // alten wert aus letztem eintrag holen
                            string oldskipcount = DataAdapter.Instance.DataCache.Objects(this.Workspace).QuerySkipStringList[DataAdapter.Instance.DataCache.Objects(this.Workspace).QuerySkipStringList.Count - 1];
                            string[] oldtokens = _Query_GetSkipTokens(oldskipcount);
                            string[] newtokens = _Query_GetSkipTokens(newskipcount);

                            // neuen wert dazurechnen
                            if (oldtokens.Count() == newtokens.Count())
                            {
                                newskipcount = "";
                                for (int i = 0; i < oldtokens.Count(); i++)
                                {
                                    int noldtoken = Int32.Parse(oldtokens[i]);
                                    int nnewtoken = Int32.Parse(newtokens[i]);
                                    nnewtoken = nnewtoken + noldtoken;
                                    if (newskipcount.Length > 0)
                                        newskipcount = newskipcount + "|";
                                    newskipcount = newskipcount + nnewtoken.ToString();
                                }
                            }
                        }
                        DataAdapter.Instance.DataCache.Objects(this.Workspace).QuerySkipStringList.Add(newskipcount);
                        break;
                }
            }
        }

        private string[] _Query_GetSkipTokens(string queryskipstring)
        {
            string[] rettokens = new string[1];
            rettokens[0] = "0";
            if (queryskipstring.Contains("|"))
            {
                char splitter = ("|".ToCharArray())[0];
                rettokens = queryskipstring.Split(splitter);
            }
            else
                rettokens[0] = queryskipstring;

            return rettokens;
        }

        /// <summary>
        /// Checks whether the togglebutton for a detailed seach is clicked, then return the last queryvalues from the infostack
        /// Also the first of the actual query-properties from the list is saved on the stack
        /// </summary>
        /// <param name="queryvalues"></param>
        /// <returns></returns>
        private List<CSQueryProperties> CheckAndApplyPossibleQueryChronicles(List<CSQueryProperties> actualQuery)
        {
            List<CSQueryProperties> retDict = new List<CSQueryProperties>();
            bool isActive = DataAdapter.Instance.DataCache.Profile.Option_GetValueBoolean(CSEnumOptions.searchovermatchlist);
            CSQueryProperties newPropForSaving = new CSQueryProperties();
            if (isActive && (this.Workspace == CSEnumProfileWorkspace.workspace_default || this.Workspace == CSEnumProfileWorkspace.workspace_searchoverall))
            {
                // get the last query and reset the toggle-button state
                retDict.Add(DataAdapter.Instance.DataCache.Info.GetLastQueryChroniclesEntry());
                DataAdapter.Instance.DataCache.Profile.Option_SetValueBoolean(CSEnumOptions.searchovermatchlist, !isActive);
                //DataAdapter.Instance.DataCache.Info.ClearQueryChronicles();
                DataAdapter.Instance.InformObservers();

                if (retDict.Count > 0)
                {
                    // Merge the old and new conditions one for saving on the stack            
                    CSQueryToken[] newMergedTokens = new CSQueryToken[actualQuery[0].searchconditions.Length + retDict[0].searchconditions.Length];
                    Array.Copy(actualQuery[0].searchconditions, newMergedTokens, actualQuery[0].searchconditions.Length);
                    Array.Copy(retDict[0].searchconditions, 0, newMergedTokens, actualQuery[0].searchconditions.Length, retDict[0].searchconditions.Length);

                    // Merge the old and new columns one for saving on the stack            
                    string[] newMergedColumns = new string[actualQuery[0].searchcolumns.Length + retDict[0].searchcolumns.Length];
                    Array.Copy(actualQuery[0].searchcolumns, newMergedColumns, actualQuery[0].searchcolumns.Length);
                    Array.Copy(retDict[0].searchcolumns, 0, newMergedColumns, actualQuery[0].searchcolumns.Length, retDict[0].searchcolumns.Length);

                    // build a new property and lay it on the stack
                    newPropForSaving.searchconditions = newMergedTokens;
                    newPropForSaving.searchcolumns = newMergedColumns;
                }
            }else
            {
                newPropForSaving = actualQuery[0];
            }

            DataAdapter.Instance.DataCache.Info.AddNewQueryChroniclesEntry(newPropForSaving);
            return retDict;
        }

        #endregion _Query (privates)

        #region NotifyPropertyChanged*****************************************************************************************************************************

        public event PropertyChangedEventHandler PropertyChanged;

        private void NotifyPropertyChanged(string propertyName)
        {
            if (PropertyChanged != null)
            {
                if (Log.Log.IsProfileEnabled) Log.Log.ProfileStart();
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
                if (Log.Log.IsProfileEnabled) Log.Log.ProfileEnd();
            }
        }

        #endregion NotifyPropertyChanged*****************************************************************************************************************************
    }  

    // TS 31.01.13 hilfsklasse um den Local Connector anzusteuern
    public class CustomHyperlinkButton : HyperlinkButton
    {
        /// <summary>
        /// Exposes the base protected OnClick method as a public method.
        /// </summary>
        public void OnClickPublic()
        {
            OnClick();
        }
    }

    //public class CustomButton : Button
    //{
    //    public void OnClickPublic()
    //    {
    //        OnClick();
    //    }
    //}
    //CustomHyperlinkButton ccc = new CustomHyperlinkButton();
    //ccc.Click += (sender, arg) =>
    //    {
    //        System.Windows.Browser.HtmlPage.Window.Navigate(new Uri("http://192.168.100.207:8080/J2C_DMS_CServer/2Charta-DMS/download/Default/Config/Downloads/LocalConnector/2Charta_LC.application"), "_self");
    //        ccc = null;
    //    };
    //ccc.OnClickPublic();
}