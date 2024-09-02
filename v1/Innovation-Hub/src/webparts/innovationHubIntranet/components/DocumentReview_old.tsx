import * as React from "react";
import { useState, useEffect } from "react";
import * as moment from "moment";
import { Web } from "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";

import {
  DetailsList,
  DetailsListLayoutMode,
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
  Rating,
  RatingSize,
  Modal,
  IModalStyles,
  Spinner,
  IColumn,
} from "@fluentui/react";
import ReactQuill from "react-quill";
import "react-quill/dist/quill.snow.css";
import "../ExternalRef/styleSheets/Styles.css";
import { mergeStyles, mergeStyleSets } from "@fluentui/react/lib/Styling";
import styles from "./InnovationHubIntranet.module.scss";
import CustomLoader from "./CustomLoader";
import alertify from "alertifyjs";
import "alertifyjs/build/css/alertify.css";

let sortDRMaster = [];
let sortDRData = [];

const DocumentReview = (props: any) => {
  // Variable-Declaration-Section Starts
  const sharepointWeb = Web(props.URL);
  const drListName = "Review Log";

  const drAllitems = [];
  const allPeoples = props.peopleList;
  const _drColumns = [
    {
      key: "Request",
      name: "Request",
      fieldName: "Request",
      minWidth: 70,
      maxWidth: 70,
      onColumnClick: (ev, column) => {
        _onColumnClick(ev, column);
      },
    },
    {
      key: "FileName",
      name: "File name",
      fieldName: "FileName",
      minWidth: 230,
      maxWidth: 230,
      onColumnClick: (ev, column) => {
        _onColumnClick(ev, column);
      },
      onRender: (item) => (
        // <div style={{ cursor: "pointer" }}>
        <TooltipHost
          id={item.ID}
          content={item.FileName}
          overflowMode={TooltipOverflowMode.Parent}
        >
          <span aria-describedby={item.ID}>{item.FileName}</span>
        </TooltipHost>
        // </div>
      ),
    },
    {
      key: "Sent",
      name: "Sent",
      fieldName: "Sent",
      minWidth: 70,
      maxWidth: 70,
      onColumnClick: (ev, column) => {
        _onColumnClick(ev, column);
      },
      onRender: (item) => moment(item.Sent).format("DD/MM/YYYY"),
    },
    {
      key: "Response",
      name: "Response",
      fieldName: "Response",
      minWidth: 80,
      maxWidth: 80,
      onColumnClick: (ev, column) => {
        _onColumnClick(ev, column);
      },
    },
    {
      key: "User",
      name: "From",
      fieldName: "User",
      minWidth: 50,
      maxWidth: 50,
      onColumnClick: (ev, column) => {
        _onColumnClick(ev, column);
      },
      onRender: (item) => (
        <>
          <div
            style={{
              display: "flex",
              alignItems: "center",
              justifyContent: "flex-start",
              cursor: "pointer",
              marginTop: -5,
            }}
          >
            <TooltipHost
              content={
                <ul style={{ margin: 10, padding: 0 }}>
                  <li>
                    <div style={{ display: "flex" }}>
                      <Persona
                        size={PersonaSize.size32}
                        presence={PersonaPresence.none}
                        imageUrl={
                          "/_layouts/15/userphoto.aspx?size=S&username=" +
                          `${item.UserDetails.UserEmail}`
                        }
                      />
                      <Label style={{ marginLeft: 10 }}>
                        {item.UserDetails.UserName}
                      </Label>
                    </div>
                  </li>
                </ul>
              }
              delay={TooltipDelay.zero}
              id={item.ID}
              directionalHint={DirectionalHint.bottomCenter}
              styles={{ root: { display: "inline-block" } }}
            >
              <Persona
                size={PersonaSize.size32}
                presence={PersonaPresence.none}
                imageUrl={
                  "/_layouts/15/userphoto.aspx?size=S&username=" +
                  `${item.UserDetails.UserEmail}`
                }
              />
            </TooltipHost>
          </div>
        </>
      ),
    },
    {
      key: "Actions",
      name: "Actions",
      fieldName: "Actions",
      minWidth: 50,
      maxWidth: 50,

      onRender: (item) => (
        <>
          <Icon
            iconName="PageArrowRight"
            className={drIconStyleClass.DetailArrow}
            onClick={async () => {
              setdrColumns(_drColumns);
              setSelectedID(item.ID);
              await setDrReviewFormDisplay({
                condition: false,
                selectedItem: {},
              });
              setDrReviewFormDisplay({
                condition: true,
                selectedItem: { ...item },
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
            }}
          />
        </>
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
    to: "Me",
    request: "All",
    response: "All",
    toUser: "",
    fromUser: "",
    fileName: "",
    product: "",
  };
  const modules = {
    toolbar: [
      [
        {
          header: [1, 2, 3, false],
        },
      ],
      ["bold", "italic", "underline"],
      [
        {
          color: [],
        },
        {
          background: [],
        },
      ],
      [
        {
          list: "ordered",
        },
        {
          list: "bullet",
        },
        {
          indent: "-1",
        },
        {
          indent: "+1",
        },
      ],
      ["clean"],
    ],
  };
  const formats = [
    "header",
    "bold",
    "italic",
    "underline",
    "list",
    "bullet",
    "indent",
    "background",
    "color",
  ];
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
      width: 670,
      selectors: {
        ".ms-DetailsRow-cell": {
          height: 40,
        },
      },
    },
    contentWrapper: {
      // height: 430,
      height: "calc(100vh - 290px)",
      overflowX: "hidden",
      overflowY: "auto",
      // overflowY: "scroll",
    },
  };
  const drDropdownStyles: Partial<IDropdownStyles> = {
    root: {
      width: 165,
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
    caretDown: { fontSize: 14, color: "#000" },
  };
  const drActiveDropdownStyles: Partial<IDropdownStyles> = {
    root: {
      width: 165,
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
    caretDown: { fontSize: 14, color: "#038387" },
  };
  const drReviewFormDropDownStyles: Partial<IDropdownStyles> = {
    root: {
      width: 190,
      margin: "10px 10px 10px 0",
    },
    title: {
      height: "36px",
      padding: "1px 10px",
    },
    caretDown: {
      fontSize: 14,
      padding: "3px",
      color: "#000",
      fontWeight: "bold",
    },
  };
  const drSearchBoxStyles: Partial<ISearchBoxStyles> = {
    root: {
      width: 165,
      marginRight: "15px",
      backgroundColor: "#F5F5F7",
      outline: "none",
      border: "1px solid #E8E8EA",
      borderRadius: "4px",
    },
    icon: { fontSize: 14, color: "#000" },
  };
  const drActiveSearchBoxStyles: Partial<ISearchBoxStyles> = {
    root: {
      width: 165,
      marginRight: "15px",
      backgroundColor: "#F5F5F7",
      outline: "none",
      border: "2px solid #038387",
      borderRadius: "4px",
    },
    field: { fontWeight: 600, color: "#038387" },
    icon: { fontSize: 14, color: "#038387" },
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
  const drReviewFormPP: Partial<IBasePickerStyles> = {
    // root: { margin: "10px 0px" },
    root: {
      width: 300,
      margin: "10px 0px",
      selectors: {
        ".ms-BasePicker-text": {
          // padding: "3px 10px",
          // border: "1px solid black",
          borderRadius: 4,
          maxHeight: 105,
          overflowX: "hidden",
          "::after": {
            border: "none",
          },
        },
      },
      ".ms-Persona-primaryText": {
        fontWeight: 600,
        border: "none",
      },
    },
    input: {
      height: 36,
      padding: "0px 10px !important",
    },
    itemsWrapper: {
      padding: "0px 5px !important",
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
  });
  // Styles-Section Ends
  // States-Declaration Starts
  const [drReRender, setDrReRender] = useState(true);
  const [currentUser, setCurrentUser] = useState({});
  const [drUnsortMasterData, setDrUnsortMasterData] = useState(drAllitems);
  const [drMasterData, setDrMasterData] = useState(drAllitems);
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
      .catch(drErrorFunction);
  };

  const drGetCurrentUserDetails = () => {
    sharepointWeb.currentUser
      .get()
      .then((user) => {
        let drCurrentUser = {
          Name: user.Title,
          Email: user.Email,
          Id: user.Id,
        };
        setCurrentUser(drCurrentUser);
      })
      .catch(drErrorFunction);
  };
  const drGetData = (peoples: any) => {
    sharepointWeb.lists
      .getByTitle(drListName)
      .items.select(
        "*",
        "FromUser/Title",
        "FromUser/Id",
        "FromUser/EMail",
        "ToUser/Title",
        "ToUser/Id",
        "ToUser/EMail",
        "CcEmail/Title",
        "CcEmail/Id",
        "CcEmail/EMail"
      )
      .expand("FromUser", "CcEmail", "ToUser")
      .top(5000)
      .orderBy("auditSent", false)
      .get()
      .then((items) => {
        items.forEach((item) => {
          let tempCcEmails = [];
          if (item.CcEmailId) {
            peoples.forEach((people) => {
              item.CcEmail.forEach((email) => {
                if (people.ID == email.Id) {
                  tempCcEmails.push(people);
                }
              });
            });
          }

          drAllitems.push({
            ID: item.Id ? item.Id : "",
            Link: item.auditLink ? item.auditLink : "",
            Request: item.auditRequestType ? item.auditRequestType : "",
            FileName: item.Title ? item.Title : "",
            Sent: item.auditSent,
            Response: item.auditResponseType ? item.auditResponseType : "",
            UserDetails: {
              UserName: item.FromUser ? item.FromUser.Title : "",
              UserEmail: item.FromUser ? item.FromUser.EMail : "",
              UserId: item.FromUser ? item.FromUser.Id : "",
            },
            User: item.FromUser ? item.FromUser.Title : "",
            ToUserDetails: {
              UserName: item.ToUser ? item.ToUser.Title : "",
              UserEmail: item.ToUser ? item.ToUser.EMail : "",
              UserId: item.ToUser ? item.ToUser.Id : "",
            },
            ToUser: item.ToUser ? item.ToUser.Title : "",
            RequestComments: item.auditComments ? item.auditComments : "",
            ResponseComments: item.Response_x0020_Comments
              ? item.Response_x0020_Comments
              : "",
            CcEmailIds: item.CcEmailId ? item.CcEmailId : [],
            CcEmails: item.CcEmailId ? tempCcEmails : [],
            Product: item.ProductName ? item.ProductName : "",
            RepeatedIssue: item.FeedbackRepeated
              ? item.FeedbackRepeated
              : false,
            Rating: item.Rating ? item.Rating : 0,
            Created: item.Created,
            Modified: item.Modified,
          });
        });

        let drAllitemsAfterInitialFilter = drAllitems.filter((item) => {
          return (
            item.Response == "Pending" &&
            item.ToUserDetails.UserName ==
              props.spcontext.pageContext.user.displayName
          );
        });

        setDrMasterData([...drAllitems]);
        sortDRMaster = drAllitems;
        setDrUnsortMasterData([...drAllitems]);
        let top500DrData = drAllitemsAfterInitialFilter.splice(0, 500);
        setDrData([...top500DrData]);
        sortDRData = [...top500DrData];
        setDrLoader("noLoader");
      })
      .catch(drErrorFunction);
  };
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
      .catch(drErrorFunction);

    //Response Choices
    sharepointWeb.lists
      .getByTitle(drListName)
      .fields.getByInternalNameOrTitle("auditResponseType")()
      .then((response) => {
        response["Choices"].forEach((choice) => {
          if (choice != null) {
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
      .catch(drErrorFunction);
  };
  const drhandleFilters = async (key: string, option: string) => {
    let _filters = { ...drFilters };
    _filters[key] = option;

    await filterItems(_filters, drMasterData);
    await setDrReviewFormDisplay({ condition: false, selectedItem: {} });
  };
  const filterItems = async (filterKeys: any, data: any) => {
    // let dataToBeFiltered = [...drMasterData];
    let dataToBeFiltered = [...data];

    if (filterKeys.view && filterKeys.view != "All") {
      if (filterKeys.view == "Pending") {
        dataToBeFiltered = dataToBeFiltered.filter((arr) => {
          return arr.Response == "Pending";
        });
      }
      if (filterKeys.view == "Pending edit") {
        dataToBeFiltered = dataToBeFiltered.filter((arr) => {
          return (
            arr.Response == "Pending" &&
            (arr.Request == "Initial Edit" || arr.Request == "Final Edit")
          );
        });
      }
      if (filterKeys.view == "Send by me") {
        dataToBeFiltered = dataToBeFiltered.filter((arr) => {
          return arr.UserDetails.UserId == currentUser["Id"];
        });
      }
      if (filterKeys.view == "Responded by me") {
        dataToBeFiltered = dataToBeFiltered.filter((arr) => {
          return (
            arr.ToUserDetails.UserId == currentUser["Id"] &&
            arr.Response != "Pending"
          );
        });
      }
      if (filterKeys.view == "Last 30 days") {
        let todayDate = moment().format("YYYY-MM-DD");
        let last30Days = moment().subtract(30, "days").format("YYYY-MM-DD");
        dataToBeFiltered = dataToBeFiltered.filter((arr) => {
          return (
            moment(arr.Sent).format("YYYY-MM-DD") >= last30Days &&
            moment(arr.Sent).format("YYYY-MM-DD") <= todayDate
          );
        });
      }
    }

    if (filterKeys.to && filterKeys.to != "Anyone") {
      if (filterKeys.to == "Me") {
        dataToBeFiltered = dataToBeFiltered.filter((arr) => {
          return arr.ToUserDetails.UserId == currentUser["Id"];
        });
      }
      if (filterKeys.to == "Me or Me Cc'd") {
        dataToBeFiltered = dataToBeFiltered.filter((arr) => {
          return (
            arr.ToUserDetails.UserId == currentUser["Id"] ||
            arr.CcEmails.some((people) => {
              people.ID == currentUser["Id"];
            }) == true
          );
        });
      }
    }

    if (filterKeys.request && filterKeys.request != "All") {
      dataToBeFiltered = dataToBeFiltered.filter((arr) => {
        return arr.Request == filterKeys.request;
      });
    }

    if (filterKeys.response && filterKeys.response != "All") {
      dataToBeFiltered = dataToBeFiltered.filter((arr) => {
        return arr.Response == filterKeys.response;
      });
    }

    if (filterKeys.toUser) {
      dataToBeFiltered = dataToBeFiltered.filter((arr) => {
        return arr.ToUserDetails.UserName.toLowerCase().includes(
          filterKeys.toUser.toLowerCase()
        );
      });
    }

    if (filterKeys.fromUser) {
      dataToBeFiltered = dataToBeFiltered.filter((arr) => {
        return arr.UserDetails.UserName.toLowerCase().includes(
          filterKeys.fromUser.toLowerCase()
        );
      });
    }

    if (filterKeys.fileName) {
      dataToBeFiltered = dataToBeFiltered.filter((arr) => {
        return arr.FileName.toLowerCase().includes(
          filterKeys.fileName.toLowerCase()
        );
      });
    }

    if (filterKeys.product) {
      dataToBeFiltered = dataToBeFiltered.filter((arr) => {
        return arr.Product.toLowerCase().includes(
          filterKeys.product.toLowerCase()
        );
      });
    }

    let top500DrData = dataToBeFiltered.splice(0, 500);
    setDrData([...top500DrData]);
    sortDRData = [...top500DrData];
    await setDrFilters({ ...filterKeys });
  };
  const GetUserDetails = (filterText: any) => {
    var result = allPeoples.filter(
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
          auditComments: `Reallocated from ${item.FromUser.Title} by ${
            currentUser["Name"]
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
        await setDrReRender(!drReRender);
        await setDrLoader("noLoader");
        ReallocatePopup();
      })
      .catch(drErrorFunction);
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
    await setDrReRender(!drReRender);
    await setDrLoader("noLoader");
    CancelRequestPopup();
  };
  const drSubmitFunction = async () => {
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
      FeedbackRepeated: drReviewFormOptionDisplay.issuesCategory.issueRepeated,
    };

    await drUpdateItem(
      requestUpdateData,
      drReviewFormDisplay.selectedItem["ID"]
    );

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
    await setDrReviewFormDisplay({
      condition: false,
      selectedItem: {},
    });
    await setDrReRender(!drReRender);
    await setDrLoader("noLoader");
    SubmitPopup();
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
            // ActivityPlanner: item.ActivityPlanner ? item.ActivityPlanner : null,
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
    await setDrReRender(!drReRender);
    await setDrLoader("noLoader");
    SignOffPopup();
  };
  const drFixLinkFunction = async () => {
    const requestUpdateData = {
      FixLink: true,
    };
    await drUpdateItem(
      requestUpdateData,
      drReviewFormDisplay.selectedItem["ID"]
    );
    await setDrLoader("noLoader");
    FixLinkPopup();
  };
  const drCreateItem = async (_createData: any) => {
    await sharepointWeb.lists
      .getByTitle("Review Log")
      .items.add(_createData)
      .then(async () => {
        await [];
      })
      .catch(drErrorFunction);
  };
  const drUpdateItem = async (_updateData: any, targetId: number) => {
    await sharepointWeb.lists
      .getByTitle("Review Log")
      .items.getById(targetId)
      .update(_updateData)
      .then(async () => {
        await [];
      })
      .catch(drErrorFunction);
  };
  const drReallocateHandler = async (key: string, option: any) => {
    let tempReallocateData = { ...drReallocateDetails };
    tempReallocateData[`${key}`] = option;
    setDrReallocateDetails({ ...tempReallocateData });
  };
  const drReviewFormOptionHandler = (optionType: string, option: any) => {
    if (optionType == "ResponseComments") {
      let tempSelectedItem = { ...drReviewFormDisplay };
      tempSelectedItem.selectedItem[optionType] = option;
      setDrReviewFormDisplay(tempSelectedItem);
    } else if (optionType == "CcEmails") {
      let tempSelectedItem = { ...drReviewFormDisplay };
      tempSelectedItem.selectedItem[optionType] = [...option];
      setDrReviewFormDisplay(tempSelectedItem);
    } else if (optionType == "rating") {
      let reviewFormOptions = { ...drReviewFormOptionDisplay };
      reviewFormOptions[optionType] = option;
      setDrReviewFormOptionDisplay({ ...reviewFormOptions });
    } else {
      let reviewFormOptions = { ...drReviewFormOptionDisplay };
      reviewFormOptions.issuesCategory[optionType] = option;
      setDrReviewFormOptionDisplay({ ...reviewFormOptions });
    }
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
    console.log(drData);
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

  {
    /*const SubmitPopup = () => (
     <MessageBar messageBarType={MessageBarType.success} isMultiline={false}>
      Document review has been successfully submitted !!!
     </MessageBar>
   );*/
  }

  const SubmitPopup = () => (
    alertify.set("notifier", "position", "top-right"),
    alertify.success(" Document is successfully submitted !!!")
  );
  const ReallocatePopup = () => (
    alertify.set("notifier", "position", "top-right"),
    alertify.success("Document is successfully reallocated !!!")
  );
  const CancelRequestPopup = () => (
    alertify.set("notifier", "position", "top-right"),
    alertify.success("Document is successfully cancelled !!!")
  );
  const SignOffPopup = () => (
    alertify.set("notifier", "position", "top-right"),
    alertify.success("Document is successfully signed off !!!")
  );
  const FixLinkPopup = () => (
    alertify.set("notifier", "position", "top-right"),
    alertify.success(" We are working on it, check back later !!!")
  );
  const ErrorPopup = () => (
    alertify.set("notifier", "position", "top-right"),
    alertify.error("Something when error, please contact system admin.")
  );

  const drErrorFunction = (error: any) => {
    console.log(error);
    ErrorPopup();
  };
  //Function-Section Ends

  useEffect(() => {
    setDrLoader("startUpLoader");
    getAllDRAdmins();
    drGetCurrentUserDetails();
    drGetData(allPeoples);
    drGetAllOptions();
  }, [drReRender]);

  return (
    <>
      <div style={{ padding: "5px 10px" }}>
        {drLoader == "startUpLoader" ? <CustomLoader /> : null}
        {/* Popup-Section Starts */}
        <div></div>
        {/* Popup-Section Ends */}
        {/* Header-Section Starts */}
        <div>
          <div
            className={styles.dpTitle}
            style={{
              justifyContent: "flex-start",
              alignItems: "flex-start",
              // marginBottom: "20px",
              marginBottom: "10px",
            }}
          >
            <Label style={{ color: "#000", fontSize: 24, padding: 0 }}>
              Document review
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
                  //display: "flex", alignItems: "center"
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
              <Label
                style={{
                  //marginLeft: 20,
                  //marginTop: 5,
                  fontSize: "13px",
                  color: "#323130",
                }}
              >
                Number of records :{" "}
                <b style={{ color: "#038387" }}>{drData.length}</b>
              </Label>
            </div>
            <div
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
            </div>
          </div>
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
              <SearchBox
                styles={
                  drFilters.toUser ? drActiveSearchBoxStyles : drSearchBoxStyles
                }
                value={drFilters.toUser}
                onChange={(e, value) => {
                  drhandleFilters("toUser", value);
                }}
              />
            </div>
            <div>
              <Label styles={drLabelStyles}>From user</Label>
              <SearchBox
                styles={
                  drFilters.fromUser
                    ? drActiveSearchBoxStyles
                    : drSearchBoxStyles
                }
                value={drFilters.fromUser}
                onChange={(e, value) => {
                  drhandleFilters("fromUser", value);
                }}
              />
            </div>
            <div>
              <Label styles={drLabelStyles}>File name</Label>
              <SearchBox
                styles={
                  drFilters.fileName
                    ? drActiveSearchBoxStyles
                    : drSearchBoxStyles
                }
                value={drFilters.fileName}
                onChange={(e, value) => {
                  drhandleFilters("fileName", value);
                }}
              />
            </div>
            <div>
              <Label styles={drLabelStyles}>Product</Label>
              <SearchBox
                styles={
                  drFilters.product
                    ? drActiveSearchBoxStyles
                    : drSearchBoxStyles
                }
                value={drFilters.product}
                onChange={(e, value) => {
                  drhandleFilters("product", value);
                }}
              />
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
                  filterItems(tempResetFilters, drUnsortMasterData);
                }}
              />
            </div>
          </div>
        </div>
        {/* Filter-Section Ends */}
        {/* Body-Section Starts */}
        <div style={{ display: "flex" }}>
          {/* DetailList-Section Starts */}
          {drData.length > 0 ? (
            <div>
              <DetailsList
                items={...drData}
                columns={drReviewFormDisplay.condition ? _drColumns : drColumns}
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
          ) : (
            <Label
              style={{
                paddingLeft: 745,
                paddingTop: 40,
              }}
              className={generalStyles.inputLabel}
            >
              No Data Found !!!
            </Label>
          )}
          {/* DetailList-Section Ends */}
          {/* Form-Section Starts */}
          <div style={{ width: "100%" }}>
            {drReviewFormDisplay.condition ? (
              <div
                style={{
                  // width: 800,
                  // height: 460,
                  marginTop: 16,
                  overflowX: "hidden",
                  overflowY: "auto",
                  maxHeight: "calc(100vh - 290px)",
                }}
                className={styles.requestReviewPanel}
              >
                <div
                  style={{
                    // height: 500,
                    // width: 750,
                    marginLeft: 20,
                    marginRight: 5,
                  }}
                >
                  <div
                    style={{
                      display: "flex",
                      justifyContent: "center",
                      alignItems: "center",
                      position: "relative",
                      width: "762px",
                    }}
                  >
                    <span className={generalStyles.titleLabel}>
                      {`Request to 
                      ${drReviewFormDisplay.selectedItem[
                        "Request"
                      ].toLowerCase()}`}
                    </span>
                    <span
                      style={{
                        color: "#959595",
                        position: "absolute",
                        right: "10px",
                        top: "0",
                        fontWeight: "500",
                      }}
                    >
                      {moment(drReviewFormDisplay.selectedItem["Sent"]).format(
                        "DD/MM/YYYY h:mm A"
                      )}
                    </span>
                  </div>
                  <div className={styles.drRequestFormBtnSection}>
                    <div style={{ display: "flex" }}>
                      <a
                        href={`${drReviewFormDisplay.selectedItem["Link"]}?web=1`}
                        data-interception="off"
                        target="_blank"
                      >
                        <button
                          // style={{
                          //   height: 37,
                          // }}
                          className={styles.openFileBtn}
                        >
                          Open file
                        </button>
                      </a>
                      <button
                        className={styles.OpenHistoryBtn}
                        onClick={() => {
                          props.handleclick(
                            "DocumentReviewHistory",
                            drReviewFormDisplay.selectedItem["ID"]
                          );
                        }}
                      >
                        History
                      </button>
                      <button
                        className={styles.fixLinkBtn}
                        onClick={() => {
                          drFixLinkFunction();
                          setDrLoader("FixLink");
                        }}
                      >
                        {drLoader == "FixLink" ? <Spinner /> : "Fix link"}
                      </button>
                    </div>
                    <div style={{ display: "flex" }}>
                      <button
                        className={
                          drReviewFormOptionDisplay.selectedOption
                            ? styles.drRequestFormSubmitBtn
                            : styles.drRequestFormBtnDisabled
                        }
                        onClick={() => {
                          if (drReviewFormOptionDisplay.selectedOption) {
                            drSubmitFunction();
                            setDrLoader("Submit");
                          }
                        }}
                      >
                        {drLoader == "Submit" ? <Spinner /> : "Submit"}
                      </button>
                      <button
                        className={
                          drReviewFormDisplay.selectedItem["Response"] ==
                          "Pending"
                            ? styles.reallocateBtn
                            : styles.disableBtn
                        }
                        onClick={() => {
                          if (
                            drReviewFormDisplay.selectedItem["Response"] ==
                            "Pending"
                          ) {
                            setDrReallocatePopup({
                              condition: true,
                              allocatedUser:
                                drReviewFormDisplay.selectedItem["UserDetails"]
                                  .UserId,
                            });
                            drReallocateHandler(
                              "reallocateUser",
                              allPeoples.filter((people) => {
                                return (
                                  people.ID ==
                                  drReviewFormDisplay.selectedItem[
                                    "UserDetails"
                                  ].UserId
                                );
                              })[0]
                            );
                          }
                        }}
                      >
                        Reallocate
                      </button>
                      <button
                        className={
                          drReviewFormDisplay.selectedItem["Response"] ==
                          "Pending"
                            ? styles.cancelRequestBtn
                            : styles.disableBtn
                        }
                        onClick={() => {
                          drReviewFormDisplay.selectedItem["Response"] ==
                          "Pending"
                            ? setDrCancelRequestPopup(true)
                            : "";
                        }}
                      >
                        {/* Cancel request */}
                        Cancel
                      </button>
                    </div>
                  </div>
                  <div
                    style={{
                      display: "flex",
                      justifyContent: "flex-start",
                      //justifyContent: "space-between"
                    }}
                  >
                    <div
                      style={{
                        width: 307,
                        paddingRight: 5,
                      }}
                      className={generalStyles.inputField}
                    >
                      <label className={generalStyles.inputLabel}>File</label>
                      <span className={generalStyles.inputValue}>
                        {drReviewFormDisplay.selectedItem["FileName"]}
                      </span>
                    </div>
                    <div
                      style={{
                        width: 325,
                        paddingRight: 5,
                      }}
                      className={generalStyles.inputField}
                    >
                      <label className={generalStyles.inputLabel}>
                        From user
                      </label>
                      <span className={generalStyles.inputValue}>
                        {
                          drReviewFormDisplay.selectedItem["UserDetails"]
                            .UserName
                        }
                      </span>
                    </div>
                    <div className={generalStyles.inputField}>
                      <label className={generalStyles.inputLabel}>
                        Current response
                      </label>
                      <span className={generalStyles.inputValue}>
                        {drReviewFormDisplay.selectedItem["Response"]}
                      </span>
                    </div>
                  </div>
                  <div
                    style={{
                      display: "flex",
                      alignItems: "baseline",
                      justifyContent: "flex-start",
                    }}
                  >
                    <div style={{ width: 293 }}>
                      <label className={generalStyles.inputLabel}>
                        CC email
                      </label>
                      <NormalPeoplePicker
                        disabled={
                          drReviewFormDisplay.selectedItem["Response"] ==
                          "Pending"
                            ? false
                            : true
                        }
                        inputProps={{
                          placeholder:
                            drReviewFormDisplay.selectedItem["Response"] ==
                            "Pending"
                              ? "Find People"
                              : "",
                        }}
                        styles={drReviewFormPP}
                        onResolveSuggestions={GetUserDetails}
                        selectedItems={
                          drReviewFormDisplay.selectedItem["CcEmails"]
                        }
                        onChange={(selectedUser) => {
                          drReviewFormOptionHandler("CcEmails", selectedUser);
                        }}
                      />
                    </div>
                    {drReviewFormDisplay.selectedItem["Response"] !=
                    "Pending" ? (
                      <div style={{ marginLeft: 20 }}>
                        <label className={generalStyles.inputLabel}>
                          Rating
                        </label>
                        <Rating
                          max={4}
                          rating={drReviewFormDisplay.selectedItem["Rating"]}
                          disabled={true}
                          style={{ width: 120 }}
                          size={RatingSize.Large}
                        />
                      </div>
                    ) : (
                      ""
                    )}
                    {drReviewFormDisplay.selectedItem["Response"] !=
                    "Pending" ? (
                      <div>
                        <label className={generalStyles.inputLabel}>
                          Repeated Issues
                        </label>
                        <Toggle
                          onText="Yes"
                          offText="No"
                          checked={
                            drReviewFormDisplay.selectedItem["RepeatedIssue"]
                          }
                          disabled={true}
                          styles={{ root: { marginTop: "15px" } }}
                        />
                      </div>
                    ) : (
                      ""
                    )}

                    {drReviewFormDisplay.selectedItem["Response"] ==
                    "Pending" ? (
                      <div
                        className={generalStyles.inputField}
                        style={{
                          display: "flex",
                          alignItems: "center",
                          justifyContent: "flex-start",
                        }}
                      >
                        <div style={{ marginLeft: "20px" }}>
                          <label className={generalStyles.inputLabel}>
                            Response
                          </label>
                          <Dropdown
                            placeholder="Select your response"
                            selectedKey={
                              drReviewFormOptionDisplay.selectedOption
                            }
                            options={
                              drReviewFormDisplay.selectedItem["Request"] ==
                              "Report"
                                ? [
                                    { key: "Select", text: "Select" },
                                    { key: "Completed", text: "Completed" },
                                    {
                                      key: "Publish ready",
                                      text: "Publish ready",
                                    },
                                  ]
                                : drReviewFormDisplay.selectedItem["Request"] ==
                                  "Review"
                                ? [
                                    { key: "Select", text: "Select" },
                                    { key: "Feedback", text: "Feedback" },
                                    { key: "Returned", text: "Returned" },
                                    { key: "Endorsed", text: "Endorsed" },
                                    { key: "Signed Off", text: "Signed Off" },
                                    {
                                      key: "Publish ready",
                                      text: "Publish ready",
                                    },
                                  ]
                                : drReviewFormDisplay.selectedItem["Request"] ==
                                  "Initial Edit"
                                ? [
                                    { key: "Select", text: "Select" },
                                    { key: "Edited", text: "Edited" },
                                    { key: "Returned", text: "Returned" },
                                    {
                                      key: "Minor feedback",
                                      text: "Minor feedback",
                                    },
                                    {
                                      key: "Major feedback",
                                      text: "Major feedback",
                                    },
                                    {
                                      key: "Publish ready",
                                      text: "Publish ready",
                                    },
                                  ]
                                : drReviewFormDisplay.selectedItem["Request"] ==
                                  "Assemble"
                                ? [
                                    { key: "Select", text: "Select" },
                                    { key: "Assembled", text: "Assembled" },
                                    { key: "Returned", text: "Returned" },
                                    {
                                      key: "Publish ready",
                                      text: "Publish ready",
                                    },
                                  ]
                                : drReviewFormDisplay.selectedItem["Request"] ==
                                  "Add Images"
                                ? [
                                    { key: "Select", text: "Select" },
                                    { key: "Inserted", text: "Inserted" },
                                    { key: "Returned", text: "Returned" },
                                    {
                                      key: "Publish ready",
                                      text: "Publish ready",
                                    },
                                  ]
                                : drReviewFormDisplay.selectedItem["Request"] ==
                                  "Publish"
                                ? [
                                    { key: "Select", text: "Select" },
                                    { key: "Published", text: "Published" },
                                    { key: "Returned", text: "Returned" },
                                    {
                                      key: "Publish ready",
                                      text: "Publish ready",
                                    },
                                    {
                                      key: "Signed Off",
                                      text: "Signed Off",
                                    },
                                  ]
                                : drReviewFormDisplay.selectedItem["Request"] ==
                                  "Final Edit"
                                ? [
                                    { key: "Select", text: "Select" },
                                    { key: "Edited", text: "Edited" },
                                    { key: "Returned", text: "Returned" },
                                    {
                                      key: "Minor feedback",
                                      text: "Minor feedback",
                                    },
                                    {
                                      key: "Major feedback",
                                      text: "Major feedback",
                                    },
                                    {
                                      key: "Publish ready",
                                      text: "Publish ready",
                                    },
                                  ]
                                : drReviewFormDisplay.selectedItem["Request"] ==
                                  "Sign-off"
                                ? [
                                    { key: "Select", text: "Select" },
                                    { key: "Signed Off", text: "Signed Off" },
                                    { key: "Feedback", text: "Feedback" },
                                    {
                                      key: "Publish ready",
                                      text: "Publish ready",
                                    },
                                  ]
                                : [{ key: "Select", text: "Select" }]
                            }
                            dropdownWidth={"auto"}
                            styles={drReviewFormDropDownStyles}
                            onChange={(e, option) => {
                              option.key != "Select"
                                ? setDrReviewFormOptionDisplay({
                                    condition: true,
                                    selectedOption: option.key,
                                    issuesCategory: {
                                      issues: "",
                                      issuesSeverity: "",
                                      issueRepeated: false,
                                    },
                                    rating: 3,
                                  })
                                : setDrReviewFormOptionDisplay({
                                    condition: false,
                                    selectedOption: null,
                                    issuesCategory: {
                                      issues: "",
                                      issuesSeverity: "",
                                      issueRepeated: false,
                                    },
                                    rating: 0,
                                  });
                            }}
                          />
                        </div>

                        <div
                          style={{
                            display: "flex",
                            alignItems: "flex-start",
                            justifyContent: "flex-start",
                          }}
                        >
                          {drReviewFormOptionDisplay.selectedOption ? (
                            <div style={{ marginLeft: "10px" }}>
                              <label
                                className={generalStyles.inputLabel}
                                style={{ marginTop: "-2px" }}
                              >
                                Ratings
                              </label>
                              <div
                                style={{
                                  display: "flex",
                                  flexDirection: "column",
                                  justifyContent: "flex-start",
                                }}
                              >
                                <Rating
                                  max={4}
                                  rating={drReviewFormOptionDisplay.rating}
                                  styles={
                                    drReviewFormOptionDisplay.rating == 4
                                      ? {
                                          ratingStarFront: { color: "#00a300" },
                                          ratingButton: { padding: "2px 2px" },
                                        }
                                      : drReviewFormOptionDisplay.rating == 3
                                      ? {
                                          ratingStarFront: { color: "#a3a300" },
                                          ratingButton: { padding: "2px 2px" },
                                        }
                                      : drReviewFormOptionDisplay.rating == 2
                                      ? {
                                          ratingStarFront: { color: "#D18700" },
                                          ratingButton: { padding: "2px 2px" },
                                        }
                                      : {
                                          ratingStarFront: { color: "#D10000" },
                                          ratingButton: { padding: "2px 2px" },
                                        }
                                  }
                                  style={{ width: 100, height: "20px" }}
                                  size={RatingSize.Large}
                                  onChange={(e, value) => {
                                    drReviewFormOptionHandler("rating", value);
                                  }}
                                />
                                <Label
                                  style={{
                                    fontSize: 13,
                                  }}
                                  styles={
                                    drReviewFormOptionDisplay.rating == 4
                                      ? { root: { color: "#00a300" } }
                                      : drReviewFormOptionDisplay.rating == 3
                                      ? { root: { color: "#a3a300" } }
                                      : drReviewFormOptionDisplay.rating == 2
                                      ? { root: { color: "#D18700" } }
                                      : { root: { color: "#D10000" } }
                                  }
                                >
                                  {drReviewFormOptionDisplay.rating == 4
                                    ? " - Exceeds"
                                    : drReviewFormOptionDisplay.rating == 3
                                    ? " - Achieved"
                                    : drReviewFormOptionDisplay.rating == 2
                                    ? " - Developing"
                                    : " - Needs improvement"}
                                </Label>
                              </div>
                            </div>
                          ) : (
                            ""
                          )}

                          {drReviewFormOptionDisplay.condition ? (
                            <div
                              className={generalStyles.inputField}
                              style={{ marginLeft: "20px", marginTop: "-4px" }}
                            >
                              <label className={generalStyles.inputLabel}>
                                Repeated issues
                              </label>
                              <Toggle
                                onText="Yes"
                                offText="No"
                                styles={{ root: { marginTop: "15px" } }}
                                onChange={(ev) => {
                                  drReviewFormOptionHandler(
                                    "issueRepeated",
                                    !drReviewFormOptionDisplay.issuesCategory
                                      .issueRepeated
                                  );
                                }}
                              />
                            </div>
                          ) : (
                            ""
                          )}
                        </div>
                      </div>
                    ) : (
                      ""
                    )}
                  </div>

                  <div>
                    <label className={generalStyles.inputLabel}>
                      Request comments
                    </label>
                    <div className={styles.reviewDesc} style={{}}>
                      {drReviewFormDisplay.selectedItem["RequestComments"]}
                    </div>
                  </div>

                  <div style={{ marginBottom: "40px" }}>
                    <label
                      className={generalStyles.inputLabel}
                      style={{ margin: "10px 10px 10px 0" }}
                    >
                      Response comments
                    </label>
                    <ReactQuill
                      theme="snow"
                      modules={modules}
                      formats={formats}
                      readOnly={
                        drReviewFormDisplay.selectedItem["Response"] ==
                        "Pending"
                          ? false
                          : true
                      }
                      value={
                        drReviewFormDisplay.selectedItem["ResponseComments"]
                          ? drReviewFormDisplay.selectedItem["ResponseComments"]
                          : ""
                      }
                      onChange={(e) => {
                        drReviewFormOptionHandler("ResponseComments", e);
                      }}
                      style={{
                        height: "auto",
                        width: "762px",
                      }}
                    ></ReactQuill>
                  </div>
                  <div>
                    <div
                      className={`${styles.drReviewSubmitBtnSection} ${generalStyles.inputField}`}
                    >
                      {(drReviewFormOptionDisplay.selectedOption == "Edited" ||
                        drReviewFormOptionDisplay.selectedOption ==
                          "Signed Off" ||
                        drReviewFormOptionDisplay.selectedOption ==
                          "Inserted") &&
                      (documentReviewAdmins.some(
                        (admin) =>
                          admin.text.toLowerCase() ==
                          currentUser["Name"].toLowerCase()
                      ) == true ||
                        currentUser["Email"] ==
                          "nprince@goodtogreatschools.org.au") ? (
                        <button
                          className={styles.drRequestFormPublishBtn}
                          onClick={() => {
                            setDrSignOffPopup(true);
                          }}
                        >
                          Sign Off and Publish
                        </button>
                      ) : (
                        ""
                      )}
                    </div>
                  </div>
                </div>
              </div>
            ) : (
              <>
                {drData.length > 0 ? (
                  <div style={{ marginLeft: 360, marginTop: 250 }}>
                    <label
                      style={{
                        color: "#959595 ",
                        display: "block",
                        fontWeight: "500",
                        margin: "5px 0",
                      }}
                    >
                      No Item Selected !!!
                    </label>
                  </div>
                ) : (
                  ""
                )}
              </>
            )}
          </div>
          {/* Form-Section Ends */}
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
          ) : (
            ""
          )}
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
          ) : (
            ""
          )}
          {/* Popup-Section Ends */}
        </div>
        {/* Body-Section Ends */}
      </div>
    </>
  );
};

export default DocumentReview;
