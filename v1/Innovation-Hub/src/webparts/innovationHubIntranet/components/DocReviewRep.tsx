import * as React from "react";
import { useState, useEffect } from "react";
import * as moment from "moment";
import { Web } from "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import * as Excel from "exceljs/dist/exceljs.min.js";
import * as FileSaver from "file-saver";
import {
  DetailsList,
  IDetailsListStyles,
  SelectionMode,
  Icon,
  Label,
  ILabelStyles,
  Toggle,
  SearchBox,
  ISearchBoxStyles,
  Dropdown,
  IDropdownStyles,
  NormalPeoplePicker,
  IBasePickerStyles,
  Persona,
  PersonaPresence,
  PersonaSize,
  TextField,
  ITextFieldStyles,
  TooltipHost,
  TooltipOverflowMode,
  TooltipDelay,
  DirectionalHint,
  Modal,
  IModalStyles,
  Spinner,
  IColumn,
} from "@fluentui/react";
import "react-quill/dist/quill.snow.css";
import "../ExternalRef/styleSheets/Styles.css";
import { mergeStyles, mergeStyleSets } from "@fluentui/react/lib/Styling";
import styles from "./InnovationHubIntranet.module.scss";
import CustomLoader from "./CustomLoader";
import alertify from "alertifyjs";
import "alertifyjs/build/css/alertify.css";
import DistributionList from "./DistributionList";
import { WebPartContext } from "@microsoft/sp-webpart-base";

import Service from "./Services";

let sortDRMaster = [];
let sortDRData = [];

let globalDLItem = null;

let arrDatas = [];

interface IProps {
  context: WebPartContext;
  spcontext: any;
  graphContent: any;
  URL: string;
  handleclick: any;
  peopleList: any;
  isAdmin: boolean;
}

interface IDocReviewDetails {
  condition: boolean;
  response: string;
  dlID: number;
  docReviewID: number;
  isEditable: boolean;
}

const DocumentReview = (props: IProps) => {
  // Variable-Declaration-Section Starts
  const sharepointWeb = Web(props.URL);
  const drListName = "Review Log";
  const dlListName: string = "Distribution List";
  let loggeduseremail: string = props.spcontext.pageContext.user.email;
  // let loggeduseremail: string = "nduff@goodtogreatschools.org.au";
  let drAllitems = [];
  const allPeoples = props.peopleList;
  let loggeduserid = allPeoples.filter(
    (dev) => dev.secondaryText == loggeduseremail
  )[0].ID;
  let loggerusername = allPeoples.filter(
    (dev) => dev.secondaryText == loggeduseremail
  )[0].text;
  let prevFirstName = null;
  let prevLasttName = null;
  const _drColumns = [
    {
      key: "FirstName",
      name: "First name",
      fieldName: "FirstName",
      minWidth: 100,
      maxWidth: 100,
      onColumnClick: (ev, column) => {
        _onColumnClick(ev, column);
      },
      onRender: (item) => {
        const showUser = prevFirstName !== item.user;
        prevFirstName = item.user;
        return (
          showUser && <TooltipHost
            id={item.ID}
            content={item.FirstName}
            overflowMode={TooltipOverflowMode.Parent}
          >
            <span aria-describedby={item.ID}>{item.FirstName}</span>
          </TooltipHost>)
      }
    },
    {
      key: "LastName",
      name: "Last name",
      fieldName: "LastName",
      minWidth: 100,
      maxWidth: 100,
      onColumnClick: (ev, column) => {
        _onColumnClick(ev, column);
      },
      onRender: (item) => {
        const showUser = prevLasttName !== item.user;
        prevLasttName = item.user;
        return (
          showUser && <TooltipHost
            id={item.ID}
            content={item.LastName}
            overflowMode={TooltipOverflowMode.Parent}
          >
            <span aria-describedby={item.ID}>{item.LastName}</span>
          </TooltipHost>
        )
      }
    },
    {
      key: "Projects",
      name: "Projects",
      fieldName: "Projects",
      minWidth: 100,
      maxWidth: 100,
      onColumnClick: (ev, column) => {
        _onColumnClick(ev, column);
      },
      onRender: (item) => (
        <TooltipHost
          id={item.ID}
          content={item.Projects}
          overflowMode={TooltipOverflowMode.Parent}
        >
          <span aria-describedby={item.ID}>{item.Projects}</span>
        </TooltipHost>
      ),
    },
    {
      key: "Task",
      name: "Task",
      fieldName: "Task",
      minWidth: 100,
      maxWidth: 100,
      onColumnClick: (ev, column) => {
        _onColumnClick(ev, column);
      },

    },
    {
      key: "Anomalies",
      name: "Anomalies",
      fieldName: "Anomalies",
      minWidth: 100,
      maxWidth: 100,
      onColumnClick: (ev, column) => {
        _onColumnClick(ev, column);
      },
      onRender: (item) => (
        <TooltipHost
          id={item.ID}
          content={item.Anomalies}
          overflowMode={TooltipOverflowMode.Parent}
        >
          <span aria-describedby={item.ID}>{item.Anomalies}</span>
        </TooltipHost>
      ),
    },
    {
      key: "NoofRevisions",
      name: "No of Revisions",
      fieldName: "NoofRevisions",
      minWidth: 100,
      maxWidth: 100,
      onColumnClick: (ev, column) => {
        _onColumnClick(ev, column);
      },
      onRender: (item) => (
        <TooltipHost
          id={item.ID}
          content={item.NoofRevisions}
          overflowMode={TooltipOverflowMode.Parent}
        >
          <span aria-describedby={item.ID}>{item.NoofRevisions}</span>
        </TooltipHost>
      ),
    },
    {
      key: "Request",
      name: "Request",
      fieldName: "Requests",
      minWidth: 100,
      maxWidth: 100,
      onColumnClick: (ev, column) => {
        _onColumnClick(ev, column);
      },
      onRender: (item) => (
        <TooltipHost
          id={item.ID}
          content={item.Request}
          overflowMode={TooltipOverflowMode.Parent}
        >
          <span aria-describedby={item.ID}>{item.Request}</span>
        </TooltipHost>
      ),
    },
    {
      key: "RequestDate",
      name: "Request Date",
      fieldName: "RequestDate",
      minWidth: 100,
      maxWidth: 100,
      onColumnClick: (ev, column) => {
        _onColumnClick(ev, column);
      },
      onRender: (item) => (
        <TooltipHost
          id={item.ID}
          content={item.RequestDate}
          overflowMode={TooltipOverflowMode.Parent}
        >
          <span aria-describedby={item.ID}>{item.RequestDate}</span>
        </TooltipHost>
      ),
    },
    {
      key: "SentTo",
      name: "Sent To",
      fieldName: "SentTo",
      minWidth: 100,
      maxWidth: 100,
      onColumnClick: (ev, column) => {
        _onColumnClick(ev, column);
      },
      onRender: (item) => (
        <TooltipHost
          id={item.ID}
          content={item.SentTo}
          overflowMode={TooltipOverflowMode.Parent}
        >
          <span aria-describedby={item.ID}>{item.SentTo}</span>
        </TooltipHost>
      ),
    },
    {
      key: "ResponseDate",
      name: "Response Date",
      fieldName: "ResponseDate",
      minWidth: 100,
      maxWidth: 100,
      onColumnClick: (ev, column) => {
        _onColumnClick(ev, column);
      },
      onRender: (item) => (
        <TooltipHost
          id={item.ID}
          content={item.ResponseDate}
          overflowMode={TooltipOverflowMode.Parent}
        >
          <span aria-describedby={item.ID}>{item.ResponseDate}</span>
        </TooltipHost>
      ),
    },
    {
      key: "Response",
      name: "Response",
      fieldName: "Response",
      minWidth: 100,
      maxWidth: 100,
      onColumnClick: (ev, column) => {
        _onColumnClick(ev, column);
      },
      onRender: (item) => (
        <TooltipHost
          id={item.ID}
          content={item.Response}
          overflowMode={TooltipOverflowMode.Parent}
        >
          <span aria-describedby={item.ID}>{item.Response}</span>
        </TooltipHost>
      ),
    },
    {
      key: "Comments",
      name: "Comments",
      fieldName: "Comments",
      minWidth: 100,
      maxWidth: 100,
      onColumnClick: (ev, column) => {
        _onColumnClick(ev, column);
      },
      onRender: (item) => (
        <TooltipHost
          id={item.ID}
          content={item.Comments}
          overflowMode={TooltipOverflowMode.Parent}
        >
          <span aria-describedby={item.ID}>{item.Comments}</span>
        </TooltipHost>
      ),
    },

  ];

  const drDrpDwnOptns = {
    viewOptns: [
      { key: "All", text: "All" },
      { key: "Pending", text: "Pending" },
      { key: "Pending edit", text: "Pending edit" },
      { key: "Send by me", text: "Send by me" },
      { key: "Responded by me", text: "Responded by me" },
      { key: "Last 30 days", text: "Last 30 days" },
    ],
    toOptns: [
      { key: "Me", text: "Me" },
      { key: "Me or Me Cc'd", text: "Me or Me Cc'd" },
      { key: "Anyone", text: "Anyone" },
    ],
    requestOptns: [{ key: "All", text: "All" }],
    responseOptns: [{ key: "All", text: "All" }],
  };
  const filters = {
    view: "Pending",
    to: "Anyone",
    request: "All",
    response: "All",
    toUser: "",
    fromUser: "",
    fileName: "",
    product: "",

    toUserSearchFlag: false,
    fromUserSearchFlag: false,
    fileNameSearchFlag: false,
    productSearchFlag: false,

    searchFlag: false,
  };

  // Variable-Declaration-Section Ends
  // Styles-Section Starts
  const drLabelStyles: Partial<ILabelStyles> = {
    root: {
      width: 150,
      marginRight: 10,
      fontSize: "13px",
      color: "#323130",
    },
  };
  const drDetailsListStyles: Partial<IDetailsListStyles> = {
    root: {
      width: "100%",
      selectors: {
        ".ms-DetailsRow-cell": {
          height: 40,
        },
      },
    },
    contentWrapper: {
      height: "calc(100vh - 290px)",
      overflowX: "hidden",
      overflowY: "auto",
    },
  };
  const drDropdownStyles: Partial<IDropdownStyles> = {
    root: {
      width: 145,
      marginRight: "15px",
      backgroundColor: "#F5F5F7",
    },
    title: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
      color: "#7C7C7C",
      border: "1px solid #E8E8EA",
      borderRadius: "4px",
    },
    dropdownItem: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
    },
    dropdownItemSelected: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
    },
    callout: {
      maxHeight: "400px !important",
    },
    caretDown: { fontSize: 14, color: "#000" },
  };
  const drActiveDropdownStyles: Partial<IDropdownStyles> = {
    root: {
      width: 145,
      marginRight: "15px",
      backgroundColor: "#F5F5F7",
    },
    title: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
      fontWeight: 600,
      color: "#038387",
      border: "2px solid #038387",
      borderRadius: "4px",
    },
    dropdownItem: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
    },
    dropdownItemSelected: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
    },
    callout: {
      maxHeight: "400px !important",
    },
    caretDown: { fontSize: 14, color: "#038387" },
  };

  const drSearchBoxStyles: Partial<ISearchBoxStyles> = {
    root: {
      position: "relative",
      width: 145,
      marginRight: 0,
      backgroundColor: "#F5F5F7",
      outline: "none",
      border: "1px solid #E8E8EA",
      borderRadius: "4px",
      borderTopRightRadius: 0,
      borderBottomRightRadius: 0,
    },
    icon: { display: "none", fontSize: 14, color: "#000" },
    iconContainer: {
      display: "none",
    },
    clearButton: {
      height: "100%",
      backgroundColor: "rgb(245, 245, 247)",
    },
  };
  const drActiveSearchBoxStyles: Partial<ISearchBoxStyles> = {
    root: {
      position: "relative",
      width: 145,
      marginRight: 0,
      backgroundColor: "#F5F5F7",
      outline: "none",
      border: "2px solid #038387",
      borderRadius: "4px",
      borderTopRightRadius: 0,
      borderBottomRightRadius: 0,
    },
    field: { fontWeight: 600, color: "#038387" },
    icon: { display: "none", fontSize: 14, color: "#038387" },
    iconContainer: {
      display: "none",
    },
    clearButton: {
      height: "100%",
      backgroundColor: "rgb(245, 245, 247)",
    },
  };
  const drModalStyles: Partial<IModalStyles> = {
    root: { borderRadius: "none" },
    main: {
      width: 500,
      margin: 10,
      padding: "20px 10px",
      display: "flex",
      flexDirection: "column",
      alignItems: "center",
      justifyContent: "center",
    },
  };
  const drModalTextFields: Partial<ITextFieldStyles> = {
    root: { width: "93%", margin: "10px 20px" },
    fieldGroup: {
      height: 40,
    },
  };

  const drModalBoxPP: Partial<IBasePickerStyles> = {
    root: {
      width: "93%",
      margin: "10px 20px",
    },
    itemsWrapper: {
      height: "30px !important",
      width: "100% !important",
      padding: "0px 3px !important",
    },
    text: {
      height: "40px !important",
      padding: "4px 3px !important",
      width: "100% !important",
    },
  };
  const generalStyles = mergeStyleSets({
    titleLabel: {
      color: "#2392B2 !important",
      fontWeight: "500",
      fontSize: "17px",
    },
    inputLabel: {
      color: "#2392B2 !important",
      display: "block",
      fontWeight: "500",
      margin: "5px 0",
    },
    inputValue: {
      color: "#000",
      fontWeight: "500",
      fontSize: "13px",
      lineBreak: "anywhere",
    },
    inputField: {
      margin: "10px 0",
    },
    sectionHeadingLabel: {
      textAlign: "left",
      color: "#000",
      paddingLeft: 15,
      fontSize: 20,
      fontWeight: 600,
    },
    DLPopupHeading: {
      color: "#2392b2",
    },
  });
  const drIconStyleClass = mergeStyleSets({
    DetailArrow: [
      {
        fontSize: 25,
        height: 14,
        width: 17,
        color: "#038387",
        margin: "0 7px",
        cursor: "pointer",
      },
    ],
    refresh: [
      {
        color: "white",
        fontSize: "18px",
        height: 20,
        width: 22,
        cursor: "pointer",
        backgroundColor: "#038387",
        padding: 5,
        marginTop: 27,
        borderRadius: 2,
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        ":hover": {
          backgroundColor: "#025d60",
        },
      },
    ],
    searchIcon: [
      {
        color: "#adadad",
        fontSize: "18px",
        fontWeight: "bold",
        height: 20,
        width: 22,
        cursor: "pointer",
        backgroundColor: "rgb(245, 245, 247)",
        padding: 5,
        borderRadius: 4,
        border: "1px solid #adadad50",
        borderTopLeftRadius: "0",
        borderBottomLeftRadius: "0",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        marginRight: 15,
        ":hover": {
          color: "rgb(3, 131, 135)",
        },
      },
    ],
    popupCloseIcon: [
      {
        color: "#ababab",
        fontSize: 18,
        fontWeight: 600,
        height: 20,
        width: 22,
        cursor: "pointer",
        backgroundColor: "#fff",
        padding: 5,
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        ":hover": {
          color: "#f00",
        },
      },
    ],
  });
  // Styles-Section Ends
  // States-Declaration Starts
  const [drReRender, setDrReRender] = useState(true);
  const [currentUser, setCurrentUser] = useState({});
  const [drUnsortMasterData, setDrUnsortMasterData] = useState([]);
  const [drMasterData, setDrMasterData] = useState([]);
  const [drData, setDrData] = useState([]);
  const [documentReviewAdmins, setDocumentReviewAdmins] = useState([]);
  const [drDropDownOptions, setDrDropDownOptions] = useState(drDrpDwnOptns);
  const [drFilters, setDrFilters] = useState(filters);
  const [drReviewFormDisplay, setDrReviewFormDisplay] = useState({
    condition: false,
    selectedItem: {},
  });
  const [drReviewFormOptionDisplay, setDrReviewFormOptionDisplay] = useState({
    condition: false,
    selectedOption: null,
    issuesCategory: { issues: "", issuesSeverity: "", issueRepeated: false },
    rating: 4,
  });
  const [drReallocatePopup, setDrReallocatePopup] = useState({
    condition: false,
    allocatedUser: null,
  });
  const [drReallocateDetails, setDrReallocateDetails] = useState({
    reallocateUser: {},
    reallocateComment: null,
  });
  const [drCancelRequestPopup, setDrCancelRequestPopup] = useState(false);
  const [drCancelReason, setDrCancelReason] = useState("");
  const [drSignOffPopup, setDrSignOffPopup] = useState(false);
  const [drSignOffOptions, setDrSignOffOptions] = useState({
    assignTo: null,
    signOffComments: "",
    publishRequestComments: "",
  });
  const [drLoader, setDrLoader] = useState("noLoader");
  const [selectedID, setSelectedID] = useState();
  const [isNewToOld, setIsNewToOld] = useState(true);
  const [drColumns, setdrColumns] = useState(_drColumns);

  const [DLdocReviewDetails, setDLdocReviewDetails] =
    useState<IDocReviewDetails>({
      condition: false,
      response: "",
      dlID: null,
      docReviewID: null,
      isEditable: null,
    });

  const [DLPopup, setDLPopup] = useState<boolean>(false);
  let DateListFormat = "DD/MM/YYYY";
  let DatePickerFormat = "YYYY-MM-DDT14:00:00Z";
  // States-Declaration Ends
  //Function-Section Starts
  const getAllDRAdmins = () => {
    sharepointWeb.siteGroups
      .getByName("Document Review Sign Off and Publish")
      .users.get()
      .then((users) => {
        let DRAdmins = [];
        users.forEach((user) => {
          DRAdmins.push({
            key: 1,
            imageUrl:
              `/_layouts/15/userphoto.aspx?size=S&accountname=` +
              `${user.Email}`,
            text: user.Title,
            ID: user.Id,
            secondaryText: user.Email,
            isValid: true,
          });
        });
        setDocumentReviewAdmins(DRAdmins);
      })
      .catch((err) => {
        drErrorFunction(err, "getAllDRAdmins");
      });
  };
  const queryGenerator = (query) => {
    let queryStr: string = "";
    if (query.length > 1) {
      var lastTwoStatement = query.length - 2;
      queryStr = "<Where><And>";
      for (let i = 0; i < query.length - 1; i++) {
        if (i == lastTwoStatement) {
          queryStr = queryStr + query[i];
          queryStr = queryStr + query[i + 1];
          break;
        } else {
          queryStr = queryStr + query[i] + "<And>";
        }
      }
      for (let i = 0; i < query.length - 1; i++) {
        queryStr += "</And>";
      }
      queryStr += "</Where>";

      queryStr = queryStr.replace(/\n/g, "");
    } else {
      queryStr = `<Where>`;
      queryStr += query[0];
      queryStr += `</Where>`;

      queryStr = queryStr.replace(/\n/g, "");
    }

    let Filtercondition = `
          <View Scope='RecursiveAll'>
            <Query>
              <OrderBy>
                <FieldRef Name='auditSent' Ascending='FALSE'/>
              </OrderBy>
              ${queryStr ? queryStr : null}
            </Query>
            <ViewFields>
            <FieldRef Name='ID' />
              <FieldRef Name='FromUser' />
              <FieldRef Name='ToUser' />
              <FieldRef Name='CcEmail' />
              <FieldRef Name='auditLink' />
              <FieldRef Name='auditRequestType' />
              <FieldRef Name='Title' />
              <FieldRef Name='auditSent' />
              <FieldRef Name='auditResponseType' />
              <FieldRef Name='auditComments' />
              <FieldRef Name='Response_x0020_Comments' />
              <FieldRef Name='ProductName' />
              <FieldRef Name='FeedbackRepeated' />
              <FieldRef Name='Rating' />
              <FieldRef Name='DistributionListID' />
              <FieldRef Name='Created' />
              <FieldRef Name='Modified' />
            </ViewFields>
            <RowLimit Paged='TRUE'>5000</RowLimit>
          </View>`;

    getThresholddata(Filtercondition);
  };
  const drGetCurrentUserDetails = () => {
    let drCurrentUser = {
      Name: loggerusername,
      Email: loggeduseremail,
      Id: loggeduserid,
    };
    setCurrentUser(drCurrentUser);
  };
  const generateExcel = () => {
    let arrExport = drData;
    const workbook = new Excel.Workbook();
    const worksheet = workbook.addWorksheet("My Sheet");
    worksheet.columns = [
      { header: "First Name", key: "FirstName", width: 25 },
      { header: "Last Name", key: "LastName", width: 25 },
      { header: "Projects", key: "Projects", width: 40 },
      { header: "Task", key: "Task", width: 50 },
      { header: "Anomalies", key: "Anomalies", width: 25 },
      { header: "No of Revisions", key: "NoofRevisions", width: 25 },
      { header: "Request", key: "Request", width: 20 },
      {
        header: "Request Date",
        key: "RequestDate",
        width: 40,
      },
      { header: "Sent To", key: "SentTo", width: 30 },
      { header: "Response Date", key: "ResponseDate", width: 30 },
      { header: "Response", key: "Response", width: 30 },
      { header: "Comments", key: "Comments", width: 30 },

    ];
    let prevName = null;
    arrExport.forEach((item) => {
      const showUser = prevName !== item.user;
      prevName = item.user;
      worksheet.addRow({
        FirstName: showUser && (item.FirstName ? item.FirstName : "")||"",
        LastName: showUser && (item.LastName ? item.LastName : "")|| "",
        Projects: item.Projects ? item.Projects : "",
        Task: item.Task ? item.Task : "",
        Anomalies: item.Anomalies ? item.Anomalies : "",
        NoofRevisions: item.NoofRevisions ? item.NoofRevisions : "",
        Request: item.Request ? item.Request : "",
        RequestDate: item.RequestDate ? item.RequestDate : "",
        SentTo: item.SentTo ? item.SentTo : "",
        ResponseDate: item.ResponseDate ? item.ResponseDate : "",
        Response: item.Response ? item.Response : "",
        Comments: item.Comments ? item.Comments : "",

      });
    });
    [
      "A1",
      "B1",
      "C1",
      "D1",
      "E1",
      "F1",
      "G1",
      "H1",
      "I1",
      "J1",
      "K1",
      "L1",

    ].map((key) => {
      worksheet.getCell(key).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "00e8d1" },
      };
    });
    [
      "A1",
      "B1",
      "C1",
      "D1",
      "E1",
      "F1",
      "G1",
      "H1",
      "I1",
      "J1",
      "K1",
      "L1",


    ].map((key) => {
      worksheet.getCell(key).color = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFFFFF" },
      };
    });
    workbook.xlsx
      .writeBuffer()
      .then((buffer) =>
        FileSaver.saveAs(
          new Blob([buffer]),
          `DocReview-${new Date().toLocaleString()}.xlsx`
        )
      )
      .catch((err) => console.log("Error writing excel export", err));
  };
  // Function to process the data and get the desired results
  const processData = (tempArrforDR) => {
    tempArrforDR.sort((a, b) => (a.user > b.user ? 1 : -1));

    // Step 2, 3, and 4: Create the result array with the "noofRep" property
    const result = [];
    const userTitleCount = {};

    tempArrforDR.forEach((item) => {
      const key = `${item.user}-${item.title}`;
      if (!userTitleCount[key]) {
        userTitleCount[key] = 0;
      }
      userTitleCount[key]++;
    });

    tempArrforDR.forEach((item) => {
      const key = `${item.user}-${item.title}`;
      if (!userTitleCount[key]) {
        return;
      }
      result.push({ ...item, NoofRevisions: userTitleCount[key] });
      delete userTitleCount[key];

    })
    const sortedarray = result.sort((a, b) => (a.user > b.user ? 1 : -1));

    return sortedarray;
  }

  function getThresholddata(Filtercondition) {
    let tempArrforDR = [];
    arrDatas = [];

    //Title field is indexed, filter / orderBy fields should still be indexed for this endpoint
    sharepointWeb.lists
      .getByTitle("Review Log")
      .renderListDataAsStream({
        ViewXml: Filtercondition,
      })
      .then((data) => {
        arrDatas.push(...data.Row);
        console.log(data, '');

        if (arrDatas.length < 2000 && data.NextHref) {
          getPagedValues(data.NextHref, Filtercondition);
        } else {
          arrDatas.forEach((item) => {
            let tempCcEmails = [];
            if (item.CcEmail != "" && item.CcEmail != null) {
              allPeoples.forEach((people) => {
                item.CcEmail.forEach((email) => {
                  if (people.ID == email.id) {
                    tempCcEmails.push(people);
                  }
                });
              });
            }

            tempArrforDR.push({
              Projects: item.ProductName ? item.ProductName : "",
              ID: item.Id ? item.Id : "",
              Link: item.auditLink ? item.auditLink : "",
              Request: item.auditRequestType ? item.auditRequestType : "",
              Anomalies: item.Title ? item.Title : "",
              Sent: item.auditSent,
              Response: item.auditResponseType ? item.auditResponseType : "",
              UserDetails: {
                UserName: item.FromUser ? item.FromUser.title : "",
                UserEmail: item.FromUser ? item.FromUser.email : "",
                UserId: item.FromUser ? item.FromUser.id : "",
              },
              RequestDate: item?.["auditSent."]
                ? moment(item?.["auditSent"], DatePickerFormat).format(DateListFormat)
                : null,
              User: item.FromUser ? item.FromUser[0].title : "",
              SentTo: item?.ToUser[0]?.title ? item?.ToUser[0]?.title : "",
              ToUserDetails: {
                UserName: item.ToUser ? item.ToUser[0].title : "",
                UserEmail: item.ToUser ? item.ToUser[0].email : "",
                UserId: item.ToUser ? item.ToUser[0].id : "",
              },
              ResponseDate: item?.["Modified."]
                ? moment(item?.["Modified."], DatePickerFormat).format(DateListFormat)
                : null,
              Comments: item.auditComments
                ? item.auditComments
                : "",
              CcEmails: item.CcEmailId ? tempCcEmails : [],
              RepeatedIssue: item.FeedbackRepeated
                ? item.FeedbackRepeated
                : false,

              Created: item.Created,
              Modified: item.Modified,
              user: item.FromUser[0].title ? item.FromUser[0].title : "",
              title: item.Title ? item.Title : "",
              request: item.auditRequestType ? item.auditRequestType : "",
              FirstName: item.FromUser[0].title ? item.FromUser[0].title.split(' ')[0] : "",
              LastName: item.FromUser[0].title ? item.FromUser[0].title.split(' ')[1] : ""

            });
          });
          let result = processData(tempArrforDR)


          console.log(result, 'checking ');
          if (drReviewFormDisplay.condition) {
            let filterItem = tempArrforDR.filter((_arr) => {
              return _arr.ID == drReviewFormDisplay.selectedItem["ID"];
            });

            if (filterItem.length > 0) {
              setDrReviewFormDisplay({
                condition: true,
                selectedItem: { ...filterItem[0] },
              });
            } else {
              setDrReviewFormDisplay({
                condition: false,
                selectedItem: null,
              });
            }
          }

          setDrMasterData([...result]);
          sortDRMaster = result;
          setDrUnsortMasterData([...result]);
          setDrData(result);
          sortDRData = [...result];
          setDrLoader("noLoader");
        }
      })
      .catch((err) => {
        drErrorFunction(err, "getDRItems-getThresholddata");
      });

    //With Paged set to true in the RowLimit, we can query the next page
  }
  function getPagedValues(data, Filtercondition) {
    let tempArrforDR = [];
    sharepointWeb.lists
      .getByTitle("Review Log")
      .renderListDataAsStream({
        ViewXml: Filtercondition,
        Paging: data.substring(1),
      })
      .then(function (data) {
        arrDatas.push(...data.Row);
        if (data.NextHref && arrDatas.length < 2000) {
          getPagedValues(data.NextHref, Filtercondition);
        } else {
          arrDatas.forEach((item) => {
            let tempCcEmails = [];
            if (item.CcEmail != "" && item.CcEmail != null) {
              allPeoples.forEach((people) => {
                item.CcEmail.forEach((email) => {
                  if (people.ID == email.id) {
                    tempCcEmails.push(people);
                  }
                });
              });
            }
            tempArrforDR.push({
              ID: item.ID ? item.ID : "",
              Link: item.auditLink ? item.auditLink : "",
              Request: item.auditRequestType ? item.auditRequestType : "",
              FileName: item.Title ? item.Title : "",
              Sent: item["auditSent."],
              Response:
                item.auditResponseType == "Distribute pending"
                  ? "Pending"
                  : item.auditResponseType
                    ? item.auditResponseType
                    : "",
              UserDetails: {
                UserName:
                  item.FromUser != "" && item.FromUser != null
                    ? item.FromUser[0].title
                    : "",
                UserEmail:
                  item.FromUser != "" && item.FromUser != null
                    ? item.FromUser[0].email
                    : "",
                UserId:
                  item.FromUser != "" && item.FromUser != null
                    ? item.FromUser[0].id
                    : "",
              },
              User:
                item.FromUser != "" && item.FromUser != null
                  ? item.FromUser[0].title
                  : "",
              ToUserDetails: {
                UserName:
                  item.ToUser != "" && item.ToUser != null
                    ? item.ToUser[0].title
                    : "",
                UserEmail:
                  item.ToUser != "" && item.ToUser != null
                    ? item.ToUser[0].email
                    : "",
                UserId:
                  item.ToUser != "" && item.ToUser != null
                    ? item.ToUser[0].id
                    : "",
              },
              ToUser:
                item.ToUser != "" && item.ToUser != null
                  ? item.ToUser[0].title
                  : "",
              RequestComments: item.auditComments ? item.auditComments : "",
              ResponseComments: item.Response_x0020_Comments
                ? item.Response_x0020_Comments
                : "",
              CcEmails: tempCcEmails,
              Product: item.ProductName ? item.ProductName : "",
              RepeatedIssue: item.FeedbackRepeated
                ? item.FeedbackRepeated
                : false,
              Rating: item.Rating ? item.Rating : 0,
              DistributionListID: item.DistributionListID
                ? item.DistributionListID
                : null,
              Created: item["Created."],
              Modified: item["Modified."],
            });
          });

          if (drReviewFormDisplay.condition) {
            let filterItem = tempArrforDR.filter((_arr) => {
              return _arr.ID == drReviewFormDisplay.selectedItem["ID"];
            });

            if (filterItem.length > 0) {
              setDrReviewFormDisplay({
                condition: true,
                selectedItem: { ...filterItem[0] },
              });
            } else {
              setDrReviewFormDisplay({
                condition: false,
                selectedItem: null,
              });
            }
          }

          setDrMasterData([...tempArrforDR]);
          sortDRMaster = tempArrforDR;
          setDrUnsortMasterData([...tempArrforDR]);
          let top500DrData = tempArrforDR.splice(0, 500);

          setDrData([...top500DrData]);
          sortDRData = [...top500DrData];
          setDrLoader("noLoader");
        }
      })
      .catch((err) => {
        drErrorFunction(err, "getDRItems-getPagedValues");
      });
  }

  const drGetAllOptions = () => {
    const _sortFilterKeys = (a, b) => {
      if (a.text < b.text) {
        return -1;
      }
      if (a.text > b.text) {
        return 1;
      }
      return 0;
    };

    //Request Choices
    sharepointWeb.lists
      .getByTitle(drListName)
      .fields.getByInternalNameOrTitle("auditRequestType")()
      .then((response) => {
        response["Choices"].forEach((choice) => {
          if (choice != null) {
            if (
              drDrpDwnOptns.requestOptns.findIndex((requestOptn) => {
                return requestOptn.key == choice;
              }) == -1
            ) {
              drDrpDwnOptns.requestOptns.push({
                key: choice,
                text: choice,
              });
            }
          }
        });
      })
      .then(() => {
        drDrpDwnOptns.requestOptns.shift();
        drDrpDwnOptns.requestOptns.sort(_sortFilterKeys);
        drDrpDwnOptns.requestOptns.unshift({ key: "All", text: "All" });
      })
      .catch((err) => {
        drErrorFunction(err, "drGetAllOptions-Request");
      });

    //Response Choices
    sharepointWeb.lists
      .getByTitle(drListName)
      .fields.getByInternalNameOrTitle("auditResponseType")()
      .then((response) => {
        response["Choices"].forEach((choice) => {
          if (choice != null && choice != "Distribute pending") {
            if (
              drDrpDwnOptns.responseOptns.findIndex((responseOptn) => {
                return responseOptn.key == choice;
              }) == -1
            ) {
              drDrpDwnOptns.responseOptns.push({
                key: choice,
                text: choice,
              });
            }
          }
        });
      })
      .then(() => {
        drDrpDwnOptns.responseOptns.push({
          key: "Cancelled",
          text: "Cancelled",
        });
        drDrpDwnOptns.responseOptns.shift();
        drDrpDwnOptns.responseOptns.sort(_sortFilterKeys);
        drDrpDwnOptns.responseOptns.unshift({ key: "All", text: "All" });
      })
      .then(() => {
        setDrDropDownOptions(drDrpDwnOptns);
      })
      .catch((err) => {
        drErrorFunction(err, "drGetAllOptions-Response");
      });
  };
  const drhandleFilters = async (key: string, option: string | boolean) => {
    let _filters = { ...drFilters };
    _filters[key] = option;

    await setDrReviewFormDisplay({ condition: false, selectedItem: {} });

    if (
      key == "toUser" ||
      key == "fromUser" ||
      key == "fileName" ||
      key == "product"
    ) {
      await setDrFilters({ ..._filters });
    } else {
      _filters.toUserSearchFlag = _filters.toUser ? true : false;
      _filters.fromUserSearchFlag = _filters.fromUser ? true : false;
      _filters.fileNameSearchFlag = _filters.fileName ? true : false;
      _filters.productSearchFlag = _filters.product ? true : false;

      await setDrFilters({ ..._filters });
      await filterItems(_filters);
    }
  };
  const filterItems = (filterKeys: any) => {
    let query: string[] = [];

    if (filterKeys.view && filterKeys.view != "All") {
      if (filterKeys.view == "Pending") {
        query.push(`<Or>
        <Eq>
          <FieldRef Name='auditResponseType' />
          <Value Type='Choice'>Pending</Value>
        </Eq>
        <Eq>
          <FieldRef Name='auditResponseType' />
          <Value Type='Choice'>Distribute pending</Value>
        </Eq>
     </Or>`);
      }
      if (filterKeys.view == "Pending edit") {
        query.push(`<And><Eq>
          <FieldRef Name='auditResponseType' />
          <Value Type='Choice'>Pending</Value>
        </Eq>
     <Or>
        <Eq>
           <FieldRef Name='auditRequestType' />
           <Value Type='Choice'>Initial Edit</Value>
        </Eq>
        <Eq>
           <FieldRef Name='auditRequestType' />
           <Value Type='Choice'>Final Edit</Value>
        </Eq>
     </Or>
     </And>`);
      }
      if (filterKeys.view == "Send by me") {
        query.push(`<Eq>
        <FieldRef Name='FromUser'/><Value Type='Text'>${loggerusername}</Value>
    </Eq>`);
      }
      if (filterKeys.view == "Responded by me") {
        query.push(`<Eq>
            <FieldRef Name='ToUser' /><Value Type='Text'>${loggerusername}</Value>
        </Eq>`);
      }
      if (filterKeys.view == "Last 30 days") {
        let todayDate = moment().format("YYYY-MM-DD");
        let last30Days = moment().subtract(30, "days").format("YYYY-MM-DD");
        query.push(`<Geq>
        <FieldRef Name='auditSent' />
        <Value IncludeTimeValue='TRUE' Type='DateTime'>${moment(
          last30Days
        ).format("YYYY-MM-DDT00:00:00Z")}</Value>
     </Geq>`);
      }
    }

    if (filterKeys.to && filterKeys.to != "Anyone") {
      if (filterKeys.to == "Me") {
        query.push(`<Eq>
        <FieldRef Name='ToUser' /><Value Type='Text'>${loggerusername}</Value>
    </Eq>`);
      }
      if (filterKeys.to == "Me or Me Cc'd") {
        query.push(`<Eq>
        <FieldRef Name='ToUser' /><Value Type='Text'>${loggerusername}</Value>
    </Eq>`);
      }
    }

    if (filterKeys.request && filterKeys.request != "All") {
      query.push(`<Eq>
      <FieldRef Name='auditRequestType' />
      <Value Type='Choice'>${filterKeys.request}</Value>
   </Eq>`);
    }

    if (filterKeys.response && filterKeys.response != "All") {
      if (filterKeys.response == "Pending") {
        query.push(`<Or>
        <Eq>
          <FieldRef Name='auditResponseType' />
          <Value Type='Choice'>Pending</Value>
        </Eq>
        <Eq>
          <FieldRef Name='auditResponseType' />
          <Value Type='Choice'>Distribute pending</Value>
        </Eq>
     </Or>`);
      } else {
        query.push(`<Eq>
          <FieldRef Name='auditResponseType' />
          <Value Type='Choice'>${filterKeys.response}</Value>
        </Eq>`);
      }
    }

    if (filterKeys.toUserSearchFlag == true && filterKeys.toUser) {
      query.push(`<Contains>
      <FieldRef Name='auditTo' />
      <Value Type='Text'>${filterKeys.toUser}</Value>
   </Contains>`);
      filterKeys.toUserSearchFlag = false;
    }

    if (filterKeys.fromUserSearchFlag == true && filterKeys.fromUser) {
      query.push(`<Contains>
      <FieldRef Name='auditFrom' />
      <Value Type='Text'>${filterKeys.fromUser}</Value>
   </Contains>`);
      filterKeys.fromUserSearchFlag = false;
    }

    if (filterKeys.fileNameSearchFlag == true && filterKeys.fileName) {
      query.push(`<Contains>
      <FieldRef Name='Title' />
      <Value Type='Text'>${filterKeys.fileName}</Value>
   </Contains>`);
      filterKeys.fileNameSearchFlag = false;
    }

    if (filterKeys.productSearchFlag == true && filterKeys.product) {
      query.push(`<Contains>
      <FieldRef Name='ProductName' />
      <Value Type='Text'>${filterKeys.product}</Value>
   </Contains>`);
      filterKeys.productSearchFlag = false;
    }

    setDrFilters({ ...filterKeys });
    setDrLoader("startUpLoader");
    queryGenerator(query);
  };
  const GetUserDetails = (filterText, currentPersonas) => {
    let _allPeoples = allPeoples;

    _allPeoples = _allPeoples.filter((arr) => {
      return arr.text.toLowerCase().indexOf("archive") == -1;
    });

    if (currentPersonas.length > 0) {
      _allPeoples = _allPeoples.filter(
        (arr) => !currentPersonas.some((persona) => persona.ID == arr.ID)
      );
    }

    var result = _allPeoples.filter(
      (value, index, self) => index === self.findIndex((t) => t.ID === value.ID)
    );

    return result.filter((item) =>
      doesTextStartWith(item.text as string, filterText)
    );
  };
  const doesTextStartWith = (text: string, filterText: string) => {
    return text.toLowerCase().indexOf(filterText.toLowerCase()) === 0;
  };
  const drReallocateFunction = async () => {
    await sharepointWeb.lists
      .getByTitle(drListName)
      .items.getById(drReviewFormDisplay.selectedItem["ID"])
      .select(
        "*",
        "FromUser/Title",
        "FromUser/Id",
        "FromUser/EMail",
        "CcEmail/Title",
        "CcEmail/Id",
        "CcEmail/EMail"
      )
      .expand("FromUser,CcEmail")
      .get()
      .then(async (item) => {
        const requestCreateData = {
          Title: item.Title ? item.Title : "",
          auditLink: item.auditLink ? item.auditLink : "",
          auditRequestType: item.auditRequestType ? item.auditRequestType : "",
          auditSent: moment().format("MM/DD/yyyy"),
          CcEmailId: item.CcEmailId
            ? { results: item.CcEmailId }
            : { results: [] },
          auditResponseType: "Pending",
          FeedbackRepeated: false,
          auditFrom: item.auditFrom ? item.auditFrom : "",
          FromEmail: item.FromEmail ? item.FromEmail : "",
          FromUserId: item.FromUserId ? item.FromUserId : "",
          auditTo: drReallocateDetails.reallocateUser["text"]
            ? drReallocateDetails.reallocateUser["text"]
            : null,
          ToEmail: drReallocateDetails.reallocateUser["secondaryText"]
            ? drReallocateDetails.reallocateUser["secondaryText"]
            : null,
          ToUserId: drReallocateDetails.reallocateUser["ID"]
            ? drReallocateDetails.reallocateUser["ID"]
            : null,
          auditComments: `Reallocated from ${item.FromUser.Title} by ${currentUser["Name"]
            } originally sent ${moment().format("DD/MM/YY HH:mm")} AEST`,
          auditDocLink: item.auditDocLink ? item.auditDocLink : "",
          auditDepartment: item.auditDepartment ? item.auditDepartment : "",
          auditLastResponse: item.auditLastResponse
            ? item.auditLastResponse
            : "",
          auditID: item.auditID ? item.auditID : "",
          auditReplyTo: item.auditReplyTo ? item.auditReplyTo : "",
          auditResponseMeetingRequired: false,
          ResponseAcknowledged: false,
          AnnualPlanID: item.AnnualPlanID ? item.AnnualPlanID : null,
          DeliveryPlanID: item.DeliveryPlanID ? item.DeliveryPlanID : null,
          ProductionBoardID: item.ProductionBoardID
            ? item.ProductionBoardID
            : null,
          DRPageName: item.DRPageName ? item.DRPageName : null,
        };
        const requestUpdateData = {
          auditComments: `Reallocated to ${drReallocateDetails.reallocateUser["text"]} by ${currentUser["Name"]}`,
          auditResponseType: "Reallocated",
          ReallocateComments: drReallocateDetails.reallocateComment
            ? drReallocateDetails.reallocateComment
            : null,
          auditLastResponse: "Pending",
          ReallocatedTo: drReallocateDetails.reallocateUser["secondaryText"]
            ? drReallocateDetails.reallocateUser["secondaryText"]
            : null,
          FeedbackRepeated:
            drReviewFormOptionDisplay.issuesCategory.issueRepeated,
        };

        await drUpdateItem(
          requestUpdateData,
          drReviewFormDisplay.selectedItem["ID"]
        );
        await drCreateItem(requestCreateData);

        await setDrReallocateDetails({
          reallocateUser: {},
          reallocateComment: null,
        });
        await setDrReallocatePopup({
          condition: false,
          allocatedUser: null,
        });
        await setDrReviewFormDisplay({
          condition: false,
          selectedItem: {},
        });
        await setDrLoader("noLoader");
        ReallocatePopup();
      })
      .catch((err) => {
        drErrorFunction(err, "drReallocateFunction-getItem");
      });
  };
  const drCancelRequestFunction = async () => {
    const requestUpdateData = {
      auditComments: `Reason cancelled: ${drCancelReason} by ${currentUser["Name"]}`,
      auditResponseType: "Cancelled",
      FeedbackRepeated: drReviewFormOptionDisplay.issuesCategory.issueRepeated,
    };
    await drUpdateItem(
      requestUpdateData,
      drReviewFormDisplay.selectedItem["ID"]
    );

    await setDrCancelReason("");
    await setDrCancelRequestPopup(false);
    await setDrReviewFormDisplay({
      condition: false,
      selectedItem: {},
    });
    await setDrLoader("noLoader");
    CancelRequestPopup();
  };

  const drSignOffFunction = async () => {
    await sharepointWeb.lists
      .getByTitle(drListName)
      .items.getById(drReviewFormDisplay.selectedItem["ID"])
      .select(
        "*",
        "FromUser/Title",
        "FromUser/Id",
        "FromUser/EMail",
        "CcEmail/Title",
        "CcEmail/Id",
        "CcEmail/EMail"
      )
      .expand("FromUser,CcEmail")
      .get()
      .then(async (item) => {
        let tempCcEmails = [];
        if (drReviewFormDisplay.selectedItem["CcEmails"].length > 0) {
          drReviewFormDisplay.selectedItem["CcEmails"].forEach((ccEmail) => {
            tempCcEmails.push(ccEmail.ID);
          });
        }

        const requestUpdateData = {
          auditResponseType: drReviewFormOptionDisplay.selectedOption
            ? drReviewFormOptionDisplay.selectedOption
            : "",
          Response_x0020_Comments:
            drReviewFormDisplay.selectedItem["ResponseComments"],
          Rating: drReviewFormOptionDisplay.rating,
          CcEmailId:
            drReviewFormDisplay.selectedItem["CcEmails"].length > 0
              ? { results: tempCcEmails }
              : { results: [] },
          FeedbackRepeated:
            drReviewFormOptionDisplay.issuesCategory.issueRepeated,
        };

        await drUpdateItem(
          requestUpdateData,
          drReviewFormDisplay.selectedItem["ID"]
        );
        if (drSignOffOptions.signOffComments) {
          const requestCreateData = {
            Title: item.Title ? item.Title : "",
            auditLink: item.auditLink ? item.auditLink : "",
            auditRequestType: "Sign-off",
            auditSent: moment().format("MM/DD/yyyy"),
            CcEmailId: item.CcEmailId
              ? { results: item.CcEmailId }
              : { results: [] },
            auditResponseType: "Signed Off",
            auditFrom: item.auditFrom ? item.auditFrom : "",
            FromEmail: item.FromEmail ? item.FromEmail : "",
            FromUserId: item.FromUserId ? item.FromUserId : "",
            auditTo: currentUser["Name"] ? currentUser["Name"] : "",
            ToEmail: currentUser["Email"] ? currentUser["Email"] : "",
            ToUserId: currentUser["Id"] ? currentUser["Id"] : "",
            auditComments: `Sign off from Editor as Client Proxy`,
            Response_x0020_Comments: `${drSignOffOptions.signOffComments}`,
            auditDocLink: item.auditDocLink ? item.auditDocLink : "",
            auditDepartment: item.auditDepartment ? item.auditDepartment : "",
            auditLastResponse: item.auditLastResponse
              ? item.auditLastResponse
              : "",
            auditID: item.auditID ? item.auditID : "",
            auditReplyTo: item.auditReplyTo ? item.auditReplyTo : "",
            AnnualPlanID: item.AnnualPlanID ? item.AnnualPlanID : null,
            DeliveryPlanID: item.DeliveryPlanID ? item.DeliveryPlanID : null,
            ProductionBoardID: item.ProductionBoardID
              ? item.ProductionBoardID
              : null,
            DRPageName: item.DRPageName ? item.DRPageName : null,
          };
          await drCreateItem(requestCreateData);
        }
        if (drSignOffOptions.assignTo) {
          const requestCreateData = {
            Title: item.Title ? item.Title : "",
            auditLink: item.auditLink ? item.auditLink : "",
            auditRequestType: "Publish",
            auditSent: moment().format("MM/DD/yyyy"),
            CcEmailId: item.CcEmailId
              ? { results: item.CcEmailId }
              : { results: [] },
            auditResponseType: "Pending",
            auditFrom: item.auditFrom ? item.auditFrom : "",
            FromEmail: item.FromEmail ? item.FromEmail : "",
            FromUserId: item.FromUserId ? item.FromUserId : "",
            auditTo: drSignOffOptions.assignTo["text"]
              ? drSignOffOptions.assignTo["text"]
              : "",
            ToEmail: drSignOffOptions.assignTo["secondaryText"]
              ? drSignOffOptions.assignTo["secondaryText"]
              : "",
            ToUserId: drSignOffOptions.assignTo["ID"]
              ? drSignOffOptions.assignTo["ID"]
              : "",
            auditComments: `${drSignOffOptions.publishRequestComments}`,
            auditDocLink: item.auditDocLink ? item.auditDocLink : "",
            auditDepartment: item.auditDepartment ? item.auditDepartment : "",
            auditLastResponse: item.auditLastResponse
              ? item.auditLastResponse
              : "",
            auditID: item.auditID ? item.auditID : "",
            auditReplyTo: item.auditReplyTo ? item.auditReplyTo : "",
            AnnualPlanID: item.AnnualPlanID ? item.AnnualPlanID : null,
            DeliveryPlanID: item.DeliveryPlanID ? item.DeliveryPlanID : null,
            ProductionBoardID: item.ProductionBoardID
              ? item.ProductionBoardID
              : null,
            DRPageName: item.DRPageName ? item.DRPageName : null,
          };
          await drCreateItem(requestCreateData);
        }
      });

    await setDrReviewFormOptionDisplay({
      condition: false,
      selectedOption: null,
      issuesCategory: {
        issues: "",
        issuesSeverity: "",
        issueRepeated: false,
      },
      rating: 0,
    });
    await setDrSignOffOptions({
      assignTo: null,
      signOffComments: "",
      publishRequestComments: "",
    });
    await setDrSignOffPopup(false);
    await setDrReviewFormDisplay({
      condition: false,
      selectedItem: {},
    });
    await setDrLoader("noLoader");
    SignOffPopup();
  };

  const drCreateItem = async (_createData: any) => {
    await sharepointWeb.lists
      .getByTitle("Review Log")
      .items.add(_createData)
      .then(async () => {
        await [];
      })
      .catch((err) => {
        drErrorFunction(err, "drCreateItem");
      });
  };
  const drUpdateItem = async (_updateData: any, targetId: number) => {
    await sharepointWeb.lists
      .getByTitle("Review Log")
      .items.getById(targetId)
      .update(_updateData)
      .then(async () => {
        await [];
      })
      .catch((err) => {
        drErrorFunction(err, "drUpdateItem");
      });
  };

  const drReallocateHandler = async (key: string, option: any) => {
    let tempReallocateData = { ...drReallocateDetails };
    tempReallocateData[`${key}`] = option;
    setDrReallocateDetails({ ...tempReallocateData });
  };

  const drSignOffHandler = (key: string, value: any) => {
    let signOffData = { ...drSignOffOptions };
    signOffData[key] = value;
    setDrSignOffOptions(signOffData);
  };

  const _onColumnClick = (
    ev: React.MouseEvent<HTMLElement>,
    column: IColumn
  ): void => {
    setdrColumns(_drColumns);
    setSelectedID(null);
    setDrReviewFormDisplay({
      condition: false,
      selectedItem: {},
    });
    const tempapColumns = _drColumns;
    const newColumns: IColumn[] = tempapColumns.slice();
    const currColumn: IColumn = newColumns.filter(
      (currCol) => column.key === currCol.key
    )[0];
    newColumns.forEach((newCol: IColumn) => {
      if (newCol === currColumn) {
        currColumn.isSortedDescending = !currColumn.isSortedDescending;
        currColumn.isSorted = true;
      } else {
        newCol.isSorted = false;
        newCol.isSortedDescending = true;
      }
    });

    const newDRData = _copyAndSort(
      sortDRData,
      currColumn.fieldName!,
      currColumn.isSortedDescending
    );
    const newDRMaster = _copyAndSort(
      sortDRMaster,
      currColumn.fieldName!,
      currColumn.isSortedDescending
    );
    setDrData([...newDRData]);
    setDrMasterData([...newDRMaster]);
  };
  function _copyAndSort<T>(
    items: T[],
    columnKey: string,
    isSortedDescending?: boolean
  ): T[] {
    let key = columnKey as keyof T;
    return items
      .slice(0)
      .sort((a: T, b: T) =>
        (isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1
      );
  }


  const ReallocatePopup = () => {
    alertify.set("notifier", "position", "top-right"),
      alertify.success("Document is successfully reallocated !!!");
    filterItems(drFilters);
  };
  const CancelRequestPopup = () => {
    alertify.set("notifier", "position", "top-right"),
      alertify.success("Document is successfully cancelled !!!");
    filterItems(drFilters);
  };
  const SignOffPopup = () => {
    alertify.set("notifier", "position", "top-right"),
      alertify.success("Document is successfully signed off !!!");
    filterItems(drFilters);
  };

  const ErrorPopup = () => {
    alertify.set("notifier", "position", "top-right"),
      alertify.error("Something when error, please contact system admin.");
    filterItems(drFilters);
  };

  const distributionListHandler = (): void => {
    let tempDrFilters = drFilters;

    tempDrFilters.toUserSearchFlag = tempDrFilters.toUser ? true : false;
    tempDrFilters.fromUserSearchFlag = tempDrFilters.fromUser ? true : false;
    tempDrFilters.fileNameSearchFlag = tempDrFilters.fileName ? true : false;
    tempDrFilters.productSearchFlag = tempDrFilters.product ? true : false;

    setDrLoader("startUpLoader");
    filterItems(tempDrFilters);
    setDLdocReviewDetails({
      condition: false,
      response: "",
      dlID: null,
      docReviewID: null,
      isEditable: null,
    });
  };

  const drErrorFunction = (error: any, functionName: string) => {
    console.log(error, functionName);

    let response = {
      ComponentName: "Review log",
      FunctionName: functionName,
      ErrorMessage: JSON.stringify(error["message"]),
      Title: loggeduseremail,
    };

    Service.SPAddItem({ Listname: "Error Log", RequestJSON: response }).then(
      () => {
        ErrorPopup();
      }
    );
  };
  //Function-Section Ends

  useEffect(() => {
    setDrLoader("startUpLoader");
    getAllDRAdmins();
    drGetCurrentUserDetails();
    filterItems(drFilters);
    drGetAllOptions();
  }, [drReRender]);

  return DLdocReviewDetails.condition ? (
    <>
      <DistributionList
        context={props.context}
        spcontext={props.spcontext}
        graphContent={props.graphContent}
        URL={props.URL}
        handleclick={props.handleclick}
        peopleList={props.peopleList}
        isAdmin={props.isAdmin}
        docReviewDetails={DLdocReviewDetails}
        distributionListHandlerFunction={distributionListHandler}
      />
    </>
  ) : (
    <>
      <div style={{ padding: "5px 10px" }}>
        {drLoader == "startUpLoader" ? <CustomLoader /> : null}
        {/* Popup-Section Starts */}
        <div></div>
        {/* Popup-Section Ends */}
        {/* Header-Section Starts */}
        <div
          className={styles.dpTitle}
          style={{
            justifyContent: "space-between",
            alignItems: "flex-start",
            marginBottom: 10,
          }}
        >
          <Label style={{ color: "#000", fontSize: 24, padding: 0 }}>
           
          </Label>
          <Label
            onClick={() => {
              generateExcel();
            }}
            style={{
              backgroundColor: "#EBEBEB",
              padding: "0 15px",
              cursor: "pointer",
              fontSize: "12px",
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              borderRadius: "3px",
              color: "#1D6F42",
              height: 34,
            }}
          >
            <Icon
              style={{
                color: "#1D6F42",
              }}
              iconName="ExcelDocument"
            // className={apIconStyleClass.export}
            />
            Export as XLS
          </Label>
        </div>
        <div
          style={{
            display: "flex",
            alignItems: "center",
            justifyContent: "space-between",
            color: "#2392b2",
          }}
        >
          <div
            style={{
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
            }}
          >
            <div
              style={{
                display: "none",
              }}
            >
              <button
                style={{
                  backgroundColor: "#038387",
                  border: 0,
                  borderRadius: 10,
                  padding: "0.25rem  0.5rem",
                  marginRight: 10,
                  cursor: "pointer",
                }}
                onClick={() => {
                  setIsNewToOld(!isNewToOld);
                  let _allData = [...drData];
                  _allData.reverse();
                  setDrData(_allData);
                  sortDRData = _allData;
                }}
              >
                <Icon
                  style={{
                    color: "#fff",
                    fontSize: isNewToOld ? 20 : 16,
                    fontWeight: isNewToOld ? "bold" : "normal",
                  }}
                  iconName="SortUp"
                  onClick={() => {
                    setIsNewToOld(true);
                    let _allData = [...drData];
                    _allData.reverse();
                    setDrData(_allData);
                    sortDRData = _allData;
                  }}
                />
                <Icon
                  style={{
                    color: "#fff",
                    fontSize: isNewToOld ? 16 : 20,
                    fontWeight: isNewToOld ? "normal" : "bold",
                  }}
                  iconName="SortDown"
                />
              </button>
              <label>{isNewToOld ? "New to old" : "Old to new"}</label>
            </div>
        
          </div>
          {/* <div
              style={{
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
              }}
            >
              <Label style={{ marginRight: "10px" }}>
                {currentUser["Name"]}
              </Label>
              <Persona
                size={PersonaSize.size24}
                presence={PersonaPresence.none}
                imageUrl={
                  "/_layouts/15/userphoto.aspx?size=S&username=" +
                  `${currentUser["Email"]}`
                }
              />
            </div> */}
        </div>
        {/* Header-Section Ends */}
        {/* Filter-Section Starts */}
        <div>
          <div
            style={{
              display: "flex",
              alignItems: "center",
              justifyContent: "flex-start",
              paddingBottom: "10px",
              flexWrap: "wrap",
            }}
          >
            <div>
              <Label styles={drLabelStyles}>View</Label>
              <Dropdown
                placeholder="Select an option"
                styles={
                  drFilters.view == "All"
                    ? drDropdownStyles
                    : drActiveDropdownStyles
                }
                options={drDropDownOptions.viewOptns}
                dropdownWidth={"auto"}
                onChange={(e, option: any) => {
                  drhandleFilters("view", option["key"]);
                }}
                selectedKey={drFilters.view}
              />
            </div>
            <div>
              <Label styles={drLabelStyles}>To</Label>
              <Dropdown
                placeholder="Select an option"
                styles={
                  drFilters.to == "Anyone"
                    ? drDropdownStyles
                    : drActiveDropdownStyles
                }
                options={drDropDownOptions.toOptns}
                dropdownWidth={"auto"}
                onChange={(e, option: any) => {
                  drhandleFilters("to", option["key"]);
                }}
                selectedKey={drFilters.to}
              />
            </div>
            <div>
              <Label styles={drLabelStyles}>Request</Label>
              <Dropdown
                placeholder="Select an option"
                styles={
                  drFilters.request == "All"
                    ? drDropdownStyles
                    : drActiveDropdownStyles
                }
                options={drDropDownOptions.requestOptns}
                dropdownWidth={"auto"}
                onChange={(e, option: any) => {
                  drhandleFilters("request", option["key"]);
                }}
                selectedKey={drFilters.request}
              />
            </div>
            <div>
              <Label styles={drLabelStyles}>Response</Label>
              <Dropdown
                placeholder="Select an option"
                styles={
                  drFilters.response == "All"
                    ? drDropdownStyles
                    : drActiveDropdownStyles
                }
                options={drDropDownOptions.responseOptns}
                dropdownWidth={"auto"}
                onChange={(e, option: any) => {
                  drhandleFilters("response", option["key"]);
                }}
                selectedKey={drFilters.response}
              />
            </div>
            <div>
              <Label styles={drLabelStyles}>To user</Label>
              <div style={{ display: "flex" }}>
                <SearchBox
                  styles={
                    drFilters.toUser
                      ? drActiveSearchBoxStyles
                      : drSearchBoxStyles
                  }
                  value={drFilters.toUser}
                  onSearch={() => {
                    drFilters.toUser
                      ? drhandleFilters("toUserSearchFlag", true)
                      : null;
                  }}
                  onChange={(e, value) => {
                    if (drFilters.toUser && value == "") {
                      let _filters = { ...drFilters };
                      _filters.toUser = "";

                      _filters.toUserSearchFlag = _filters.toUser
                        ? true
                        : false;
                      _filters.fromUserSearchFlag = _filters.fromUser
                        ? true
                        : false;
                      _filters.fileNameSearchFlag = _filters.fileName
                        ? true
                        : false;
                      _filters.productSearchFlag = _filters.product
                        ? true
                        : false;

                      setDrFilters({ ..._filters });
                      filterItems(_filters);
                    } else {
                      drhandleFilters("toUser", value);
                    }
                  }}
                />
                <Icon
                  iconName="Search"
                  title="Click to search"
                  className={drIconStyleClass.searchIcon}
                  onClick={() => {
                    drFilters.toUser
                      ? drhandleFilters("toUserSearchFlag", true)
                      : null;
                  }}
                />
              </div>
            </div>
            <div>
              <Label styles={drLabelStyles}>From user</Label>
              <div style={{ display: "flex" }}>
                <SearchBox
                  styles={
                    drFilters.fromUser
                      ? drActiveSearchBoxStyles
                      : drSearchBoxStyles
                  }
                  value={drFilters.fromUser}
                  onSearch={() => {
                    drFilters.fromUser
                      ? drhandleFilters("fromUserSearchFlag", true)
                      : null;
                  }}
                  onChange={(e, value) => {
                    if (drFilters.fromUser && value == "") {
                      let _filters = { ...drFilters };
                      _filters.fromUser = "";

                      _filters.toUserSearchFlag = _filters.toUser
                        ? true
                        : false;
                      _filters.fromUserSearchFlag = _filters.fromUser
                        ? true
                        : false;
                      _filters.fileNameSearchFlag = _filters.fileName
                        ? true
                        : false;
                      _filters.productSearchFlag = _filters.product
                        ? true
                        : false;

                      setDrFilters({ ..._filters });
                      filterItems(_filters);
                    } else {
                      drhandleFilters("fromUser", value);
                    }
                  }}
                />
                <Icon
                  iconName="Search"
                  title="Click to search"
                  className={drIconStyleClass.searchIcon}
                  onClick={() => {
                    drFilters.fromUser
                      ? drhandleFilters("fromUserSearchFlag", true)
                      : null;
                  }}
                />
              </div>
            </div>
            <div>
              <Label styles={drLabelStyles}>File name</Label>
              <div style={{ display: "flex" }}>
                <SearchBox
                  styles={
                    drFilters.fileName
                      ? drActiveSearchBoxStyles
                      : drSearchBoxStyles
                  }
                  onSearch={() => {
                    drFilters.fileName
                      ? drhandleFilters("fileNameSearchFlag", true)
                      : null;
                  }}
                  value={drFilters.fileName}
                  onChange={(e, value) => {
                    if (drFilters.fileName && value == "") {
                      let _filters = { ...drFilters };
                      _filters.fileName = "";

                      _filters.toUserSearchFlag = _filters.toUser
                        ? true
                        : false;
                      _filters.fromUserSearchFlag = _filters.fromUser
                        ? true
                        : false;
                      _filters.fileNameSearchFlag = _filters.fileName
                        ? true
                        : false;
                      _filters.productSearchFlag = _filters.product
                        ? true
                        : false;

                      setDrFilters({ ..._filters });
                      filterItems(_filters);
                    } else {
                      drhandleFilters("fileName", value);
                    }
                  }}
                />
                <Icon
                  iconName="Search"
                  title="Click to search"
                  className={drIconStyleClass.searchIcon}
                  onClick={() => {
                    drFilters.fileName
                      ? drhandleFilters("fileNameSearchFlag", true)
                      : null;
                  }}
                />
              </div>
            </div>
            <div>
              <Label styles={drLabelStyles}>Product</Label>
              <div style={{ display: "flex" }}>
                <SearchBox
                  styles={
                    drFilters.product
                      ? drActiveSearchBoxStyles
                      : drSearchBoxStyles
                  }
                  onSearch={() => {
                    drFilters.product
                      ? drhandleFilters("productSearchFlag", true)
                      : null;
                  }}
                  value={drFilters.product}
                  onChange={(e, value) => {
                    if (drFilters.product && value == "") {
                      let _filters = { ...drFilters };
                      _filters.product = "";

                      _filters.toUserSearchFlag = _filters.toUser
                        ? true
                        : false;
                      _filters.fromUserSearchFlag = _filters.fromUser
                        ? true
                        : false;
                      _filters.fileNameSearchFlag = _filters.fileName
                        ? true
                        : false;
                      _filters.productSearchFlag = _filters.product
                        ? true
                        : false;

                      setDrFilters({ ..._filters });
                      filterItems(_filters);
                    } else {
                      drhandleFilters("product", value);
                    }
                  }}
                />
                <Icon
                  iconName="Search"
                  title="Click to search"
                  className={drIconStyleClass.searchIcon}
                  onClick={() => {
                    drFilters.product
                      ? drhandleFilters("productSearchFlag", true)
                      : null;
                  }}
                />
              </div>
            </div>
            <div>
              <Icon
                iconName="Refresh"
                title="Click to reset"
                className={drIconStyleClass.refresh}
                onClick={() => {
                  let tempResetFilters = {
                    view: "Pending",
                    to: "Me",
                    request: "All",
                    response: "All",
                    toUser: "",
                    fromUser: "",
                    fileName: "",
                    product: "",
                  };
                  setSelectedID(null);
                  setDrReviewFormDisplay({
                    condition: false,
                    selectedItem: {},
                  });
                  setdrColumns(_drColumns);
                  setDrMasterData([...drUnsortMasterData]);
                  sortDRMaster = drUnsortMasterData;
                  filterItems(tempResetFilters);
                }}
              />
            </div>
          </div>
        </div>
        {/* Filter-Section Ends */}
        {/* Body-Section Starts */}
        <div style={{ display: "flex", flexDirection: "column" }}>
          {/* DetailList-Section Starts */}
          {
            drData.length > 0 ? (
              <div>
                <DetailsList
                  items={...drData}
                  columns={
                    drReviewFormDisplay.condition ? _drColumns : drColumns
                  }
                  styles={drDetailsListStyles}
                  setKey="set"
                  selectionMode={SelectionMode.none}
                  data-is-scrollable={true}
                  onShouldVirtualize={() => {
                    return false;
                  }}
                  onRenderRow={(data, defaultRender) => (
                    <div className="red">
                      {defaultRender({
                        ...data,
                        styles: {
                          root: {
                            background:
                              data.item.ID == selectedID
                                ? "linear-gradient(90deg, rgba(250,163,50,0.1491947120645133) 35%, rgba(3,131,135,0.14639359161633403) 100%)"
                                : "#fff",
                            selectors: {
                              "&:hover": {
                                background:
                                  data.item.ID == selectedID
                                    ? "linear-gradient(270deg, rgba(250,163,50,0.19961488013174022) 35%, rgba(3,131,135,0.19961488013174022) 100%)"
                                    : "#f3f2f1",
                              },
                            },
                          },
                        },
                      })}
                    </div>
                  )}
                />
              </div>
            ) : null

          }
          {/* DetailList-Section Ends */}

          {/* Popup-Section Starts */}
          {drReallocatePopup.condition ? (
            <Modal
              isOpen={drReallocatePopup.condition}
              isBlocking={true}
              styles={drModalStyles}
            >
              <div>
                <Label className={styles.drPopupLabel}>Reallocate</Label>
                <div className={styles.drPopupDescription}>
                  This will close this request and create a new request to the
                  selected user
                </div>
                <div>
                  <NormalPeoplePicker
                    styles={drModalBoxPP}
                    onResolveSuggestions={GetUserDetails}
                    itemLimit={1}
                    defaultSelectedItems={allPeoples.filter((people) => {
                      return people.ID == drReallocatePopup.allocatedUser;
                    })}
                    onChange={(selectedUser) => {
                      drReallocateHandler("reallocateUser", selectedUser[0]);
                    }}
                  />
                  <TextField
                    styles={drModalTextFields}
                    onChange={(e, value) => {
                      drReallocateHandler("reallocateComment", value);
                    }}
                    placeholder="Reason to reallocate"
                  />
                </div>
                <div className={styles.drPopupButtonSection}>
                  <button
                    className={
                      drReallocateDetails.reallocateUser
                        ? styles.successBtnActive
                        : styles.successBtnInActive
                    }
                    onClick={() => {
                      if (drReallocateDetails.reallocateUser) {
                        setDrLoader("Reallocate");
                        drReallocateFunction();
                      }
                    }}
                  >
                    {drLoader == "Reallocate" ? <Spinner /> : "Reallocate"}
                  </button>
                  <button
                    className={styles.closeBtn}
                    onClick={() => {
                      setDrReallocatePopup({
                        condition: false,
                        allocatedUser: null,
                      });
                      setDrReallocateDetails({
                        reallocateUser: {},
                        reallocateComment: null,
                      });
                    }}
                  >
                    Close
                  </button>
                </div>
              </div>
            </Modal>
          ) : (
            ""
          )}
          {drCancelRequestPopup ? (
            <Modal
              isOpen={drCancelRequestPopup}
              isBlocking={true}
              styles={drModalStyles}
            >
              <div>
                <Label className={styles.drPopupLabel}>Confirmation</Label>
                <div className={styles.drPopupDescription}>
                  This will cancel the request and remove from your review log.
                  Kindly mention the reason to cancel.
                </div>
                <TextField
                  styles={drModalTextFields}
                  onChange={(e, value) => {
                    setDrCancelReason(value);
                  }}
                  placeholder="Reason for cancelling"
                />
                <div className={styles.drPopupDescription}>
                  Do you wish to proceed?
                </div>
                <div className={styles.drPopupButtonSection}>
                  <button
                    className={
                      drCancelReason
                        ? styles.successBtnActive
                        : styles.successBtnInActive
                    }
                    onClick={() => {
                      if (drCancelReason) {
                        setDrLoader("cancelRequest");
                        drCancelRequestFunction();
                      }
                    }}
                  >
                    {drLoader == "cancelRequest" ? <Spinner /> : "Yes"}
                  </button>
                  <button
                    className={styles.closeBtn}
                    onClick={() => {
                      setDrCancelRequestPopup(false);
                      setDrCancelReason("");
                    }}
                  >
                    No
                  </button>
                </div>
              </div>
            </Modal>
          ) : null}
          {drSignOffPopup ? (
            <Modal
              isOpen={drSignOffPopup}
              isBlocking={true}
              styles={drModalStyles}
            >
              <div>
                <Label className={styles.drPopupLabel}>
                  Sign Off and Publish
                </Label>
                <div className={styles.drPopupDescription}>
                  This will save your current response and then sign off and
                  publish (if selected)
                </div>
                <NormalPeoplePicker
                  inputProps={{
                    placeholder:
                      "Assign Publish to, leave blank to not publish",
                  }}
                  styles={drModalBoxPP}
                  onResolveSuggestions={GetUserDetails}
                  itemLimit={1}
                  onChange={(selectedUser) => {
                    drSignOffHandler("assignTo", selectedUser[0]);
                  }}
                />
                <TextField
                  styles={drModalTextFields}
                  defaultValue={drSignOffOptions.signOffComments}
                  onChange={(e, value) => {
                    drSignOffHandler("signOffComments", value);
                  }}
                  placeholder="Sign Off Comments"
                />
                <TextField 
                  styles={drModalTextFields}
                  defaultValue={drSignOffOptions.publishRequestComments}
                  onChange={(e, value) => {
                    drSignOffHandler("publishRequestComments", value);
                  }}
                  placeholder="Publish request comments (if publishing)"
                />
                <div className={styles.drPopupButtonSection}>
                  <button
                    className={
                      drSignOffOptions.assignTo ||
                        drSignOffOptions.signOffComments ||
                        drSignOffOptions.publishRequestComments
                        ? styles.successBtnActive
                        : styles.successBtnInActive
                    }
                    onClick={() => {
                      if (
                        drSignOffOptions.assignTo ||
                        drSignOffOptions.signOffComments ||
                        drSignOffOptions.publishRequestComments
                      ) {
                        setDrLoader("signOff");
                        drSignOffFunction();
                      }
                    }}
                  >
                    {drLoader == "signOff" ? <Spinner /> : "Yes"}
                  </button>
                  <button
                    className={styles.closeBtn}
                    onClick={() => {
                      setDrSignOffPopup(false);
                      setDrSignOffOptions({
                        assignTo: null,
                        signOffComments: "",
                        publishRequestComments: "",
                      });
                    }}
                  >
                    Close
                  </button>
                </div>
              </div>
            </Modal>
          ) : null}
          {DLPopup ? (
            <Modal isOpen={DLPopup} isBlocking={true}>
              <div style={{ padding: "20px 40px" }}>
                <div
                  style={{
                    display: "flex",
                    alignItems: "center",
                    justifyContent: "center",
                  }}
                >
                  <Label
                    className={styles.atpPopupLabel}
                    style={{ marginLeft: "auto" }}
                  >
                    Distribute
                  </Label>
                  <div
                    className={styles.DR_DLPopupBtnSection}
                    style={{ marginLeft: "auto" }}
                  >
                    <div>
                      <Icon
                        iconName="ChromeClose"
                        title="Close"
                        className={drIconStyleClass.popupCloseIcon}
                        onClick={() => {
                          globalDLItem = null;
                          setDLPopup(false);
                        }}
                      />
                    </div>
                  </div>
                </div>

                <div className={styles.atpPopupInnerSection}>
                  {/* GGSA */}
                  <div style={{ marginBottom: 20 }}>
                    <Label className={generalStyles.sectionHeadingLabel}>
                      GGSA
                    </Label>
                    <div
                      style={{
                        display: "flex",
                        justifyContent: "flex-start",
                        margin: 15,
                        marginBottom: 10,
                      }}
                    >
                      <div
                        style={{
                          width: 400,
                          wordBreak: "break-all",
                          paddingRight: 20,
                        }}
                      >
                        <Label className={generalStyles.DLPopupHeading}>
                          Innovation team
                        </Label>
                        <>{globalDLItem["InnovationTeam"]}</>
                      </div>
                      <div
                        style={{
                          width: 400,
                          wordBreak: "break-all",
                          paddingRight: 20,
                        }}
                      >
                        <Label className={generalStyles.DLPopupHeading}>
                          Employees
                        </Label>
                        <>{globalDLItem["Employees"]}</>
                      </div>
                      <div style={{ width: 400, wordBreak: "break-all" }}>
                        <Label className={generalStyles.DLPopupHeading}>
                          Contractors
                        </Label>
                        <>{globalDLItem["Contractors"]}</>
                      </div>
                    </div>
                    <div
                      style={{
                        display: "flex",
                        justifyContent: "flex-start",
                        margin: 15,
                        marginBottom: 10,
                      }}
                    >
                      <div
                        style={{
                          width: 400,
                          wordBreak: "break-all",
                          paddingRight: 20,
                        }}
                      >
                        <Label className={generalStyles.DLPopupHeading}>
                          Business area
                        </Label>
                        <>
                          {globalDLItem["BusinessArea"]
                            ? globalDLItem["BusinessArea"].join(", ")
                            : null}
                        </>
                      </div>
                      <div
                        style={{
                          width: 400,
                          wordBreak: "break-all",
                          paddingRight: "20px",
                        }}
                      >
                        <Label className={generalStyles.DLPopupHeading}>
                          ManagementTeam
                        </Label>
                        <>{globalDLItem["ManagementTeam"]}</>
                      </div>
                      <div style={{ width: 400, wordBreak: "break-all" }}>
                        <Label className={generalStyles.DLPopupHeading}>
                          Board
                        </Label>
                        <>{globalDLItem["Board"]}</>
                      </div>
                    </div>
                    <div
                      style={{
                        display: "flex",
                        justifyContent: "flex-start",
                        margin: "15px",
                        marginBottom: "10px",
                      }}
                    >
                      <div
                        style={{
                          width: 400,
                          wordBreak: "break-all",
                          paddingRight: "20px",
                        }}
                      >
                        <Label className={generalStyles.DLPopupHeading}>
                          Comments
                        </Label>
                        <>{globalDLItem["ggsaComments"]}</>
                      </div>
                    </div>
                  </div>

                  {/* Schools */}
                  <div style={{ marginBottom: 20 }}>
                    <Label className={generalStyles.sectionHeadingLabel}>
                      Schools
                    </Label>
                    <div
                      style={{
                        display: "flex",
                        justifyContent: "flex-start",
                        margin: "15px",
                        marginBottom: "10px",
                      }}
                    >
                      <div
                        style={{
                          width: 400,
                          wordBreak: "break-all",
                          paddingRight: "20px",
                        }}
                      >
                        <Label className={generalStyles.DLPopupHeading}>
                          Principals
                        </Label>
                        <>{globalDLItem["Principals"]}</>
                      </div>
                      <div
                        style={{
                          width: 400,
                          wordBreak: "break-all",
                          paddingRight: "20px",
                        }}
                      >
                        <Label className={generalStyles.DLPopupHeading}>
                          Instruction coaches
                        </Label>
                        <>{globalDLItem["InstructionCoaches"]}</>
                      </div>
                      <div style={{ width: 400, wordBreak: "break-all" }}>
                        <Label className={generalStyles.DLPopupHeading}>
                          Teachers
                        </Label>
                        <>{globalDLItem["Teachers"]}</>
                      </div>
                    </div>
                    <div
                      style={{
                        display: "flex",
                        justifyContent: "flex-start",
                        margin: "15px",
                        marginBottom: "10px",
                      }}
                    >
                      <div
                        style={{
                          width: 400,
                          wordBreak: "break-all",
                          paddingRight: "20px",
                        }}
                      >
                        <Label className={generalStyles.DLPopupHeading}>
                          Teaching assistants
                        </Label>
                        <>{globalDLItem["TeachingAssistant"]}</>
                      </div>
                      <div
                        style={{
                          width: 400,
                          wordBreak: "break-all",
                          paddingRight: "20px",
                        }}
                      >
                        <Label className={generalStyles.DLPopupHeading}>
                          Schools
                        </Label>
                        <>{globalDLItem["Schools"]}</>
                      </div>
                      <div style={{ width: 400, wordBreak: "break-all" }}>
                        <Label className={generalStyles.DLPopupHeading}>
                          School teams
                        </Label>
                        <>{globalDLItem["SchoolTeams"]}</>
                      </div>
                    </div>
                    <div
                      style={{
                        display: "flex",
                        justifyContent: "flex-start",
                        margin: "15px",
                        marginBottom: "10px",
                      }}
                    >
                      <div
                        style={{
                          width: 400,
                          wordBreak: "break-all",
                          paddingRight: "20px",
                        }}
                      >
                        <Label className={generalStyles.DLPopupHeading}>
                          Parents
                        </Label>
                        <>{globalDLItem["Parents"]}</>
                      </div>
                      <div
                        style={{
                          width: 400,
                          wordBreak: "break-all",
                          paddingRight: "20px",
                        }}
                      >
                        <Label className={generalStyles.DLPopupHeading}>
                          Communities
                        </Label>
                        <>{globalDLItem["Communities"]}</>
                      </div>
                      <div style={{ width: 400, wordBreak: "break-all" }}>
                        <Label className={generalStyles.DLPopupHeading}>
                          Head of departments
                        </Label>
                        <>{globalDLItem["HODs"]}</>
                      </div>
                    </div>
                    <div
                      style={{
                        display: "flex",
                        justifyContent: "flex-start",
                        margin: "15px",
                        marginBottom: "10px",
                      }}
                    >
                      <div
                        style={{
                          width: 400,
                          wordBreak: "break-all",
                          paddingRight: "20px",
                        }}
                      >
                        <Label className={generalStyles.DLPopupHeading}>
                          Others
                        </Label>
                        <>{globalDLItem["SchoolOthers"]}</>
                      </div>
                      <div
                        style={{
                          width: 400,
                          wordBreak: "break-all",
                          paddingRight: "20px",
                        }}
                      >
                        <Label className={generalStyles.DLPopupHeading}>
                          Comments
                        </Label>
                        <>{globalDLItem["SchoolsComments"]}</>
                      </div>
                    </div>
                  </div>

                  {/* Digital Channels */}
                  <div style={{ marginBottom: 20 }}>
                    <Label className={generalStyles.sectionHeadingLabel}>
                      Digital channels
                    </Label>
                    <div
                      style={{
                        display: "flex",
                        justifyContent: "flex-start",
                        margin: "15px",
                        marginBottom: "10px",
                      }}
                    >
                      <div
                        style={{
                          width: 400,
                          wordBreak: "break-all",
                          paddingRight: "20px",
                        }}
                      >
                        <Label className={generalStyles.DLPopupHeading}>
                          Great teaching portal
                        </Label>
                        <>{globalDLItem["GreatTeachingPortal"]}</>
                      </div>
                      <div
                        style={{
                          width: 400,
                          wordBreak: "break-all",
                          paddingRight: "20px",
                        }}
                      >
                        <Label className={generalStyles.DLPopupHeading}>
                          Digital database
                        </Label>
                        <>{globalDLItem["DigitalDatabase"]}</>
                      </div>
                      <div style={{ width: 400, wordBreak: "break-all" }}>
                        <Label className={generalStyles.DLPopupHeading}>
                          Website
                        </Label>
                        <>{globalDLItem["Website"]}</>
                      </div>
                    </div>
                    <div
                      style={{
                        display: "flex",
                        justifyContent: "flex-start",
                        margin: "15px",
                        marginBottom: "10px",
                      }}
                    >
                      <div
                        style={{
                          width: 400,
                          wordBreak: "break-all",
                          paddingRight: "20px",
                        }}
                      >
                        <Label className={generalStyles.DLPopupHeading}>
                          Intranet
                        </Label>
                        <>{globalDLItem["Intranet"]}</>
                      </div>
                      <div
                        style={{
                          width: 400,
                          wordBreak: "break-all",
                          paddingRight: "20px",
                        }}
                      >
                        <Label className={generalStyles.DLPopupHeading}>
                          Social media (Facebook ,LinkedIn, Twitter)
                        </Label>
                        <>{globalDLItem["SocialMedia"]}</>
                      </div>
                      <div style={{ width: 400, wordBreak: "break-all" }}>
                        <Label className={generalStyles.DLPopupHeading}>
                          CEO linkedin
                        </Label>
                        <>{globalDLItem["CEOLinkedIn"]}</>
                      </div>
                    </div>
                    <div
                      style={{
                        display: "flex",
                        justifyContent: "flex-start",
                        margin: "15px",
                        marginBottom: "10px",
                      }}
                    >
                      <div
                        style={{
                          width: 400,
                          wordBreak: "break-all",
                          paddingRight: "20px",
                        }}
                      >
                        <Label className={generalStyles.DLPopupHeading}>
                          Products
                        </Label>
                        <>{globalDLItem["Products"]}</>
                      </div>
                      <div
                        style={{
                          width: 400,
                          wordBreak: "break-all",
                          paddingRight: "20px",
                        }}
                      >
                        <Label className={generalStyles.DLPopupHeading}>
                          Others
                        </Label>
                        <>{globalDLItem["DCOthers"]}</>
                      </div>
                      <div style={{ width: 400, wordBreak: "break-all" }}>
                        <Label className={generalStyles.DLPopupHeading}>
                          Collateral cabinet
                        </Label>
                        <>{globalDLItem["CollateralCabinet"]}</>
                      </div>
                    </div>
                    <div
                      style={{
                        display: "flex",
                        justifyContent: "flex-start",
                        margin: "15px",
                        marginBottom: "10px",
                      }}
                    >
                      <div
                        style={{
                          width: 400,
                          wordBreak: "break-all",
                          paddingRight: "20px",
                        }}
                      >
                        <Label className={generalStyles.DLPopupHeading}>
                          Brochure stand
                        </Label>
                        <>{globalDLItem["BrochureStand"]}</>
                      </div>
                      <div
                        style={{
                          width: 400,
                          wordBreak: "break-all",
                          paddingRight: "20px",
                        }}
                      >
                        <Label className={generalStyles.DLPopupHeading}>
                          Comments
                        </Label>
                        <>{globalDLItem["DCComments"]}</>
                      </div>
                    </div>
                  </div>

                  {/* Partners */}
                  <div style={{ marginBottom: 20 }}>
                    <Label className={generalStyles.sectionHeadingLabel}>
                      Parterns
                    </Label>
                    <div
                      style={{
                        display: "flex",
                        justifyContent: "flex-start",
                        margin: "15px",
                        marginBottom: "10px",
                      }}
                    >
                      <div
                        style={{
                          width: 400,
                          wordBreak: "break-all",
                          paddingRight: "20px",
                        }}
                      >
                        <Label className={generalStyles.DLPopupHeading}>
                          Professionals
                        </Label>
                        <>{globalDLItem["Professionals"]}</>
                      </div>
                      <div
                        style={{
                          width: 400,
                          wordBreak: "break-all",
                          paddingRight: "20px",
                        }}
                      >
                        <Label className={generalStyles.DLPopupHeading}>
                          Services
                        </Label>
                        <>{globalDLItem["Services"]}</>
                      </div>
                      <div style={{ width: 400, wordBreak: "break-all" }}>
                        <Label className={generalStyles.DLPopupHeading}>
                          Funding
                        </Label>
                        <>{globalDLItem["Funding"]}</>
                      </div>
                    </div>
                    <div
                      style={{
                        display: "flex",
                        justifyContent: "flex-start",
                        margin: "15px",
                        marginBottom: "10px",
                      }}
                    >
                      <div
                        style={{
                          width: 400,
                          wordBreak: "break-all",
                          paddingRight: "20px",
                        }}
                      >
                        <Label className={generalStyles.DLPopupHeading}>
                          Resources
                        </Label>
                        <>{globalDLItem["Resources"]}</>
                      </div>
                      <div
                        style={{
                          width: 400,
                          wordBreak: "break-all",
                          paddingRight: "20px",
                        }}
                      >
                        <Label className={generalStyles.DLPopupHeading}>
                          Governments
                        </Label>
                        <>{globalDLItem["Governments"]}</>
                      </div>
                      <div style={{ width: 400, wordBreak: "break-all" }}>
                        <Label className={generalStyles.DLPopupHeading}>
                          Indigenous
                        </Label>
                        <>{globalDLItem["Indigenous"]}</>
                      </div>
                    </div>
                    <div
                      style={{
                        display: "flex",
                        justifyContent: "flex-start",
                        margin: "15px",
                        marginBottom: "10px",
                      }}
                    >
                      <div
                        style={{
                          width: 400,
                          wordBreak: "break-all",
                          paddingRight: "20px",
                        }}
                      >
                        <Label className={generalStyles.DLPopupHeading}>
                          Professional associations
                        </Label>
                        <>{globalDLItem["ProfessionalAssociations"]}</>
                      </div>
                      <div
                        style={{
                          width: 400,
                          wordBreak: "break-all",
                          paddingRight: "20px",
                        }}
                      >
                        <Label className={generalStyles.DLPopupHeading}>
                          Media
                        </Label>
                        <>{globalDLItem["Media"]}</>
                      </div>
                      <div style={{ width: 400, wordBreak: "break-all" }}>
                        <Label className={generalStyles.DLPopupHeading}>
                          Comments
                        </Label>
                        <>{globalDLItem["PartnersComments"]}</>
                      </div>
                    </div>
                  </div>

                  {/* Politicians */}
                  <div style={{ marginBottom: 20 }}>
                    {globalDLItem["PoliticiansJSON"] ? (
                      <Label className={generalStyles.sectionHeadingLabel}>
                        Politicians
                      </Label>
                    ) : null}
                    {globalDLItem["PoliticiansJSON"]
                      ? JSON.parse(globalDLItem["PoliticiansJSON"]).map(
                        (arr) => {
                          return (
                            <>
                              {" "}
                              <div
                                style={{
                                  display: "flex",
                                  justifyContent: "flex-start",
                                  margin: "15px",
                                  marginBottom: "10px",
                                }}
                              >
                                {" "}
                                <div
                                  style={{
                                    width: 400,
                                    height: 200,
                                    wordBreak: "break-all",
                                    overflow: "auto",
                                    paddingRight: "20px",
                                  }}
                                >
                                  {" "}
                                  <Label
                                    className={generalStyles.DLPopupHeading}
                                  >
                                    {" "}
                                    State and territory{" "}
                                  </Label>{" "}
                                  <>{arr.SAT}</>{" "}
                                </div>{" "}
                                <div
                                  style={{
                                    width: 400,
                                    height: 200,
                                    wordBreak: "break-all",
                                    overflow: "auto",
                                    paddingRight: "20px",
                                  }}
                                >
                                  {" "}
                                  <Label
                                    className={generalStyles.DLPopupHeading}
                                  >
                                    {" "}
                                    Senate{" "}
                                  </Label>{" "}
                                  <>
                                    {" "}
                                    {arr.senate
                                      ? arr.senate.join(", ")
                                      : null}{" "}
                                  </>{" "}
                                </div>{" "}
                                <div
                                  style={{
                                    width: 400,
                                    height: 200,
                                    wordBreak: "break-all",
                                    overflow: "auto",
                                  }}
                                >
                                  {" "}
                                  <Label
                                    className={generalStyles.DLPopupHeading}
                                  >
                                    {" "}
                                    Head of representatives{" "}
                                  </Label>{" "}
                                  <>
                                    {arr.HOR ? arr.HOR.join(", ") : null}
                                  </>{" "}
                                </div>{" "}
                              </div>{" "}
                              <div
                                style={{
                                  display: "flex",
                                  justifyContent: "flex-start",
                                  margin: "15px",
                                  marginBottom: "10px",
                                }}
                              >
                                {" "}
                                <div
                                  style={{
                                    width: 400,
                                    height: 200,
                                    wordBreak: "break-all",
                                    overflow: "auto",
                                    paddingRight: "20px",
                                  }}
                                >
                                  {" "}
                                  <Label
                                    className={generalStyles.DLPopupHeading}
                                  >
                                    {" "}
                                    Comments{" "}
                                  </Label>{" "}
                                  <>{arr.Comments}</>{" "}
                                </div>{" "}
                              </div>{" "}
                            </>
                          );
                        }
                      )
                      : null}
                  </div>
                </div>

                <div className={styles.DR_DLPopupBtnSection}>
                  <button
                    onClick={() => {
                      globalDLItem = null;
                      setDLPopup(false);
                    }}
                  >
                    Ok
                  </button>
                </div>
              </div>
            </Modal>
          ) : null}
          {/* Popup-Section Ends */}
        </div>

        {drData.length == 0 ? (
          <div
            style={{
              display: "flex",
              justifyContent: "center",
              marginTop: "15px",
            }}
          >
            <Label style={{ color: "#2392B2" }}>No data Found !!!</Label>
          </div>
        ) : null}
        {/* Body-Section Ends */}
      </div>
    </>
  );
};

export default DocumentReview;
