import * as React from "react";
import { useState, useEffect } from "react";
import { Web } from "@pnp/sp/webs";
import * as moment from "moment";
import {
  DetailsList,
  DetailsListLayoutMode,
  SelectionMode,
  Icon,
  Label,
  Dropdown,
  IDropdownStyles,
  TooltipHost,
  TooltipOverflowMode,
  IColumn,
  ILabelStyles,
  Persona,
  PersonaPresence,
  PersonaSize,
  Modal,
  TextField,
  ITextFieldStyles,
  Spinner,
} from "@fluentui/react";
import "../ExternalRef/styleSheets/Styles.css";
import styles from "./InnovationHubIntranet.module.scss";
import Pagination from "office-ui-fabric-react-pagination";
import { mergeStyles, mergeStyleSets } from "@fluentui/react/lib/Styling";
import CustomLoader from "./CustomLoader";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import * as Excel from "exceljs/dist/exceljs.min.js";
import * as FileSaver from "file-saver";

import Service from "../components/Services";

import alertify from "alertifyjs";
import "alertifyjs/build/css/alertify.css";
let DateListFormat = "DD/MM/YYYY";

interface IProps {
  context: WebPartContext;
  spcontext: any;
  graphContent: any;
  URL: string;
  peopleList: any;
  isAdmin: boolean;
}
interface IFilter {
  BA: string;
  Title: string;
  Frequency: string;
  Provider: string;
}
interface IDropdown {
  key: string;
  text: string;
}
interface IDropdowns {
  BA: IDropdown[];
  Title: IDropdown[];
  Frequency: IDropdown[];
  Provider: IDropdown[];
}
interface IData {
  ID: number;
  ConfigID: number;
  BA: string;
  Title: string;
  Frequency: string;
  Provider: any;
  TimePeriod: string;
  Year: number;
  TLink: string;
  DocLink: string;
  DueDate: string;
  Status: string;
  GroupID: number;
  UserType: string[];
  ParentID: number;
  MasterData: string;
}
interface IPopup {
  condition: boolean;
  item: IData;
  Status: string;
  comments: string;
  statusValidation: boolean;
  commentsValidation: boolean;
  overAllValidation: boolean;
}

let sortORData: IData[] = [];
let sortORFilterData: IData[] = [];

let CurrentPage: number = 1;
let totalPageItems: number = 10;

const ORModalDrpDwnOptns = {
  ApproverStatus: [
    { key: "Not actioned", text: "Not actioned" },
    { key: "Actioned", text: "Actioned" },
    { key: "Endorsed", text: "Endorsed" },
    { key: "Not endorsed", text: "Not endorsed" },
  ],
  AudienceStatus: [
    { key: "Not actioned", text: "Not actioned" },
    { key: "Actioned", text: "Actioned" },
  ],
};

const OrgApprovalRequests = (props: IProps): JSX.Element => {
  // variable-Declaration Starts
  const sharepointWeb: any = Web(props.URL);
  const allPeoples: any[] = props.peopleList;

  const OrgReportListName = "OrgReporting";

  const currentLoggedUserEmail: string = props.spcontext.pageContext.user.email;
  // const currentLoggedUserEmail: string = `lally@goodtogreatschools.org.au`;
  const currentLoggedUserID: number = props.peopleList.filter((user) => {
    return user.secondaryText == currentLoggedUserEmail;
  })[0].ID;

  const OR_Year: number = moment().year();
  const OR_Week: string = `Week ${moment().isoWeek()}`;
  const OR_Month: string = `Month ${moment().format("MMMM")}`;
  const OR_Term: string =
    moment().month() + 1 >= 10
      ? "Term 4"
      : moment().month() + 1 >= 7
      ? "Term 3"
      : moment().month() + 1 >= 4
      ? "Term 2"
      : "Term 1";

  const _ORApprovalRequestsColumn: IColumn[] = [
    {
      key: "Column1",
      name: "Business area",
      fieldName: "BA",
      minWidth: 120,
      maxWidth: 120,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
      onRender: (item) =>
        BAacronymsCollection.filter((ba) => {
          return ba.Name == item.BA;
        })[0].ShortName,
    },
    {
      key: "Column2",
      name: "Title",
      fieldName: "Title",
      minWidth: 100,
      maxWidth: 300,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
      onRender: (item) => (
        <>
          <TooltipHost
            id={item.ID}
            content={item.Title}
            overflowMode={TooltipOverflowMode.Parent}
          >
            <span aria-describedby={item.ID}>{item.Title}</span>
          </TooltipHost>
        </>
      ),
    },
    {
      key: "Column3",
      name: "Frequency",
      fieldName: "Frequency",
      minWidth: 100,
      maxWidth: 200,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
    },
    {
      key: "Column4",
      name: "Due date",
      fieldName: "DueDate",
      minWidth: 100,
      maxWidth: 150,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
      onRender: (item) => {
        let frequencyType =
          item.Frequency == "Weekly"
            ? "LDW"
            : item.Frequency == "Monthly"
            ? "LDM"
            : item.Frequency == "Term"
            ? "LDT"
            : "";
        let typeAbbreviations =
          item.Frequency == "Weekly"
            ? "Last day of week"
            : item.Frequency == "Monthly"
            ? "Last day of month"
            : item.Frequency == "Term"
            ? "Last day of term"
            : "";
        return (
          <>
            <div
              style={{
                marginTop: "-6px",
                cursor: "pointer",
              }}
              title={
                moment(item.DueDate, DateListFormat).format(DateListFormat) +
                ` ( ${typeAbbreviations} )`
              }
            >
              {moment(item.DueDate, DateListFormat).format(DateListFormat) +
                ` ( ${frequencyType} )`}
            </div>
          </>
        );
      },
    },
    {
      key: "Column5",
      name: "Provider",
      fieldName: "Provider",
      minWidth: 150,
      maxWidth: 250,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
      onRender: (item) => (
        <>
          {item.Provider ? (
            <>
              {
                <div
                  style={{
                    display: "flex",
                    alignItems: "center",
                    justifyContent: "flex-start",
                    cursor: "pointer",
                  }}
                >
                  <div title={item.Provider.text} style={{ display: "flex" }}>
                    <Persona
                      showOverflowTooltip
                      size={PersonaSize.size24}
                      presence={PersonaPresence.none}
                      showInitialsUntilImageLoads={true}
                      imageUrl={
                        "/_layouts/15/userphoto.aspx?size=S&username=" +
                        `${item.Provider.secondaryText}`
                      }
                      styles={{
                        root: {
                          marginLeft: "12px",
                        },
                      }}
                    />
                  </div>
                  <div>
                    <span style={{ fontSize: "13px" }}>
                      {item.Provider.text}
                    </span>
                  </div>
                </div>
              }
            </>
          ) : (
            ""
          )}
        </>
      ),
    },
    {
      key: "Column6",
      name: "Open file",
      fieldName: "Doc Link",
      minWidth: 100,
      maxWidth: 100,
      onRender: (item) => (
        <>
          <a
            href={item.DocLink ? `${item.DocLink}?web=1` : null}
            data-interception="off"
            target="_blank"
          >
            <Icon
              iconName="NavigateExternalInline"
              title="Open document"
              className={ORiconStyleClass.fileOpenIcon}
              onClick={() => {}}
            />
          </a>
        </>
      ),
    },
    {
      key: "Column7",
      name: "Action",
      fieldName: "Action",
      minWidth: 100,
      maxWidth: 100,
      onRender: (item) => (
        <>
          <Icon
            iconName="FileComment"
            title="Open document"
            className={ORiconStyleClass.actionIcon}
            onClick={() => {
              setORPopup({
                condition: true,
                item: item,
                Status: "Not actioned",
                comments: null,
                statusValidation: false,
                commentsValidation: false,
                overAllValidation: false,
              });
            }}
          />
        </>
      ),
    },
  ];
  const BAacronymsCollection = [
    {
      Name: "PD Curriculum",
      ShortName: "PDC",
    },
    {
      Name: "PD Professional Learning",
      ShortName: "PDPL",
    },
    {
      Name: "PD School Improvements",
      ShortName: "PDSI",
    },
    {
      Name: "SS Business",
      ShortName: "SSB",
    },
    {
      Name: "SS Publishing",
      ShortName: "SSP",
    },
    {
      Name: "SS Content Creation",
      ShortName: "SSCC",
    },
    {
      Name: "SS Marketing",
      ShortName: "SSM",
    },
    {
      Name: "SS Technology",
      ShortName: "SST",
    },
    {
      Name: "SS Research and Evaluation",
      ShortName: "SSRE",
    },
    {
      Name: "SD School Partnerships",
      ShortName: "SSPSP",
    },
  ];
  const ORFilterKeys: IFilter = {
    BA: "All",
    Title: "All",
    Frequency: "All",
    Provider: "All",
  };
  const ORFilterOptns: IDropdowns = {
    BA: [{ key: "All", text: "All" }],
    Title: [{ key: "All", text: "All" }],
    Frequency: [{ key: "All", text: "All" }],
    Provider: [{ key: "All", text: "All" }],
  };
  // variable-Declaration Starts
  // Style-Section Starts
  const ORModalTxtFieldStyles: Partial<ITextFieldStyles> = {
    root: {
      width: 400,
      marginBottom: 20,
      borderRadius: 4,
    },
    field: { fontSize: 14, color: "#000", minHeight: "98px" },
    fieldGroup: {
      minHeight: "100px",
      background: "#fff",
      borderRadius: "4px",
    },
  };
  const ORModalErrorTxtFieldStyles: Partial<ITextFieldStyles> = {
    root: {
      width: 400,
      marginBottom: 20,
      borderRadius: 4,
    },
    field: { fontSize: 14, color: "#000", minHeight: "98px" },
    fieldGroup: {
      border: "2px solid #f00",
      minHeight: "100px",
      background: "#fff",
      borderRadius: "4px",
    },
  };
  const ORModalDropdownStyles: Partial<IDropdownStyles> = {
    root: {
      width: 186,
      marginRight: "15px",
      backgroundColor: "#fff",
    },
    title: {
      backgroundColor: "#fff",
      fontSize: 12,
      color: "#000",
      border: "1px solid #000",
      borderRadius: "4px",
    },
    dropdownItem: {
      backgroundColor: "#fff",
      fontSize: 12,
    },
    dropdownItemSelected: {
      backgroundColor: "#fff",
      fontSize: 12,
    },
    caretDown: { fontSize: 14, color: "#000" },
  };
  const ORModalErrorDropdownStyles: Partial<IDropdownStyles> = {
    root: {
      width: 186,
      marginRight: "15px",
      backgroundColor: "#fff",
    },
    title: {
      backgroundColor: "#fff",
      fontSize: 12,
      color: "#000",
      border: "2px solid #f00",
      borderRadius: "4px",
    },
    dropdownItem: {
      backgroundColor: "#fff",
      fontSize: 12,
    },
    dropdownItemSelected: {
      backgroundColor: "#fff",
      fontSize: 12,
    },
    caretDown: { fontSize: 14, color: "#000" },
  };
  const ORfilterLabelStyles: Partial<ILabelStyles> = {
    root: {
      width: 150,
      marginRight: 10,
      fontSize: "13px",
      color: "#323130",
    },
  };
  const ORDropdownStyles: Partial<IDropdownStyles> = {
    root: {
      width: 186,
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
  const ORActiveDropdownStyles: Partial<IDropdownStyles> = {
    root: {
      width: 186,
      marginRight: "15px",
      backgroundColor: "#F5F5F7",
    },
    title: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
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
    caretDown: { fontSize: 14, color: "#000" },
  };
  const ORiconStyleClass = mergeStyleSets({
    refresh: {
      color: "white",
      fontSize: "18px",
      height: 22,
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
    export: {
      color: "#038387",
      fontSize: "18px",
      height: 20,
      width: 20,
      cursor: "pointer",
      marginRight: 5,
    },
    fileOpenIcon: {
      color: "#038387",
      fontSize: 22,
      height: 22,
      width: 22,
      cursor: "pointer",
      marginRight: 5,
      marginLeft: "25px",
    },
    actionIcon: {
      color: "#038387",
      fontSize: 20,
      height: 20,
      width: 20,
      cursor: "pointer",
      marginRight: 5,
      marginLeft: "15px",
    },
  });
  // Style-Section Ends
  // State-Declaration Starts
  const [ORReRender, setORReRender] = useState<boolean>(false);
  const [ORMasterData, setORMasterData] = useState<IData[]>([]);
  const [ORData, setORData] = useState<IData[]>([]);
  const [ORDisplayData, setORDisplayData] = useState<IData[]>([]);
  const [ORFilter, setORFilter] = useState<IFilter>(ORFilterKeys);
  const [ORFilterData, setORFilterData] = useState<IData[]>([]);
  const [ORFilterDrpDown, setORFilterDrpDown] =
    useState<IDropdowns>(ORFilterOptns);
  const [ORColumns, setORColumns] = useState<IColumn[]>(
    _ORApprovalRequestsColumn
  );
  const [ORCurrentPage, setORCurrentPage] = useState<number>(CurrentPage);
  const [ORLoader, setORLoader] = useState("noLoader");
  const [ORPopup, setORPopup] = useState<IPopup>({
    condition: false,
    item: null,
    Status: null,
    comments: null,

    statusValidation: false,
    commentsValidation: false,

    overAllValidation: false,
  });
  // State-Declaration Ends
  // Function-Declaration Starts

  const generateExcel = (): void => {
    let arrExport = ORFilterData;
    const workbook = new Excel.Workbook();
    const worksheet = workbook.addWorksheet("My Sheet");
    worksheet.columns = [
      { header: "Business area", key: "BA", width: 25 },
      { header: "Title", key: "Title", width: 25 },
      { header: "Frequency", key: "Frequency", width: 25 },
      { header: "DueDate", key: "DueDate", width: 20 },
      { header: "Provider", key: "Provider", width: 25 },
      { header: "Open file", key: "DocLink", width: 25 },
    ];
    arrExport.forEach((item) => {
      worksheet.addRow({
        BA: item.BA ? item.BA : "",
        Title: item.Title ? item.Title : "",
        Frequency: item.Frequency ? item.Frequency : "",
        Provider: item.Provider.length > 0 ? item.Provider[0].text : "",
        DocLink: item.DocLink ? item.DocLink : "",
        DueDate: item.DueDate
          ? moment(item.DueDate, DateListFormat).format(DateListFormat)
          : "",
      });
    });
    ["A1", "B1", "C1", "D1", "E1", "F1"].map((key) => {
      worksheet.getCell(key).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "00e8d1" },
      };
    });
    ["A1", "B1", "C1", "D1", "E1", "F1"].map((key) => {
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
          `Organisationreport-Approvalrequests-${new Date().toLocaleString()}.xlsx`
        )
      )
      .catch((err) => console.log("Error writing excel export", err));
  };

  const onChangeFilter = (key: string, option: string): void => {
    let tempData: IData[] = ORData;
    let tempFilterKeys = ORFilter;
    tempFilterKeys[key] = option;
    if (tempFilterKeys.BA != "All") {
      tempData = tempData.filter((arr) => {
        return arr.BA == tempFilterKeys.BA;
      });
    }
    if (tempFilterKeys.Title != "All") {
      tempData = tempData.filter((arr) => {
        return arr.Title == tempFilterKeys.Title;
      });
    }
    if (tempFilterKeys.Frequency != "All") {
      tempData = tempData.filter((arr) => {
        return arr.Frequency == tempFilterKeys.Frequency;
      });
    }
    if (tempFilterKeys.Provider != "All") {
      tempData = tempData.filter((arr) => {
        return arr.Provider.text == tempFilterKeys.Provider;
      });
    }

    setORFilterData([...tempData]);
    sortORFilterData = tempData;
    setORFilter({ ...tempFilterKeys });
    paginateFunction(1, tempData);
  };
  const paginateFunction = (pagenumber, data): void => {
    if (data.length > 0) {
      let lastIndex: number = pagenumber * totalPageItems;
      let firstIndex: number = lastIndex - totalPageItems;
      let paginatedItems = data.slice(firstIndex, lastIndex);
      CurrentPage = pagenumber;
      setORDisplayData(paginatedItems);
      setORCurrentPage(pagenumber);
    } else {
      setORDisplayData([]);
      setORCurrentPage(1);
    }
  };
  const ErrorPopup = () => (
    alertify.set("notifier", "position", "top-right"),
    alertify.error("Something when error, please contact system admin.")
  );
  const ORErrorFunction = (error: any, functionName: string): void => {
    console.log(error, functionName);

    let response = {
      ComponentName: "Org reporting - requests",
      FunctionName: functionName,
      ErrorMessage: JSON.stringify(error["message"]),
      Title: currentLoggedUserEmail,
    };

    Service.SPAddItem({ Listname: "Error Log", RequestJSON: response }).then(
      () => {
        setORPopup({
          condition: false,
          item: null,
          Status: null,
          comments: null,

          statusValidation: false,
          commentsValidation: false,

          overAllValidation: false,
        });
        setORLoader("noLoader");
        ErrorPopup();
      }
    );
  };

  const ORDataHandler = (key: string, value: string): void => {
    if (key == "Status") {
      ORPopup.Status = value;
      ORPopup.statusValidation = false;
    } else if (key == "comments") {
      ORPopup.comments = value;
      ORPopup.commentsValidation = false;
    }
    ORPopup.overAllValidation = false;

    setORPopup({ ...ORPopup });
  };
  const ORValidationFunction = (): void => {
    if (ORPopup.Status == "Not actioned") {
      ORPopup.statusValidation = true;
      ORPopup.overAllValidation = true;
    }
    if (!ORPopup.comments) {
      ORPopup.commentsValidation = true;
      ORPopup.overAllValidation = true;
    }

    if (ORPopup.statusValidation || ORPopup.commentsValidation) {
      setORPopup({ ...ORPopup });
      setORLoader("noLoader");
    } else {
      ORUpdateOrgReportFunction();
    }
  };

  const ORUpdateOrgReportFunction = (): void => {
    if (ORPopup.item.MasterData == "Yes") {
      sharepointWeb.lists
        .getByTitle(OrgReportListName)
        .items.getById(ORPopup.item.ID)
        .get()
        .then((item) => {
          let ApproverCount: number = item.ApproverCount;
          let ApproverActioned: number = item.ApproverActioned;

          let updatedApproverActioned: number = ApproverActioned + 1;

          let responseData_ = {
            Status: ORPopup.Status,
            Comments: ORPopup.comments,
            ApproverActioned: updatedApproverActioned,
            DisplayStatus:
              ApproverCount == updatedApproverActioned
                ? "Read"
                : item.DisplayStatus,
            ActionedOn: moment().format("YYYY/MM/DD"),
          };

          sharepointWeb.lists
            .getByTitle(OrgReportListName)
            .items.getById(ORPopup.item.ID)
            .update(responseData_)
            .then((_item) => {
              setORPopup({
                condition: false,
                item: null,
                Status: null,
                comments: null,
                statusValidation: false,
                commentsValidation: false,
                overAllValidation: false,
              });
              setORLoader("noLoader");
              setORReRender(!ORReRender);
            })
            .catch((error) => {
              ORErrorFunction(error, "updateProviderData");
            });
        })
        .catch((err) => {
          ORErrorFunction(err, "getProviderData");
        });
    } else {
      let responseData = {
        Status: ORPopup.Status,
        Comments: ORPopup.comments,
      };
      sharepointWeb.lists
        .getByTitle(OrgReportListName)
        .items.getById(ORPopup.item.ID)
        .update(responseData)
        .then(() => {
          sharepointWeb.lists
            .getByTitle(OrgReportListName)
            .items.getById(ORPopup.item.ParentID)
            .get()
            .then((item) => {
              let ApproverCount: number = item.ApproverCount;
              let ApproverActioned: number = item.ApproverActioned;

              let updatedApproverActioned: number = ApproverActioned + 1;

              let responseData_ = {
                ApproverActioned: updatedApproverActioned,
                DisplayStatus:
                  ApproverCount == updatedApproverActioned
                    ? "Read"
                    : item.DisplayStatus,
              };

              sharepointWeb.lists
                .getByTitle(OrgReportListName)
                .items.getById(ORPopup.item.ParentID)
                .update(responseData_)
                .then((item) => {
                  setORPopup({
                    condition: false,
                    item: null,
                    Status: null,
                    comments: null,
                    statusValidation: false,
                    commentsValidation: false,
                    overAllValidation: false,
                  });
                  setORLoader("noLoader");
                  setORReRender(!ORReRender);
                })
                .catch((error) => {
                  ORErrorFunction(error, "UpdateProviderData");
                });
            })
            .catch((err) => {
              ORErrorFunction(err, "getProviderData");
            });
        })
        .catch((err) => {
          ORErrorFunction(err, "ORUpdateOrgReportFunction");
        });
    }
  };

  const reloadFilterDropdowns = (data: IData[]): void => {
    const sortFilterKeys = (a, b) => {
      if (a.text < b.text) {
        return -1;
      }
      if (a.text > b.text) {
        return 1;
      }
      return 0;
    };
    const usersOrderFunction = (dropDown): any => {
      let nonArchived = dropDown.filter((user) => {
        return !user.text.includes("Archive");
      });

      let archived = dropDown.filter((user) => {
        return user.text.includes("Archive");
      });

      return nonArchived.concat(archived);
    };
    data.forEach((item) => {
      if (
        ORFilterOptns.BA.findIndex((BA) => {
          return BA.key == item.BA;
        }) == -1 &&
        item.BA
      ) {
        ORFilterOptns.BA.push({
          key: item.BA,
          text: item.BA,
        });
      }
      if (
        ORFilterOptns.Title.findIndex((Title) => {
          return Title.key == item.Title;
        }) == -1 &&
        item.Title
      ) {
        ORFilterOptns.Title.push({
          key: item.Title,
          text: item.Title,
        });
      }
      if (
        ORFilterOptns.Frequency.findIndex((Frequency) => {
          return Frequency.key == item.Frequency;
        }) == -1 &&
        item.Frequency
      ) {
        ORFilterOptns.Frequency.push({
          key: item.Frequency,
          text: item.Frequency,
        });
      }
      if (
        ORFilterOptns.Provider.findIndex((ProviderOptn) => {
          return ProviderOptn.key == item.Provider.text;
        }) == -1 &&
        item.Provider
      ) {
        ORFilterOptns.Provider.push({
          key: item.Provider.text,
          text: item.Provider.text,
        });
      }
    });

    if (
      ORFilterOptns.Provider.some((ProviderOptn) => {
        return (
          ProviderOptn.text.toLowerCase() ==
          props.spcontext.pageContext.user.displayName.toLowerCase()
        );
      })
    ) {
      ORFilterOptns.Provider.shift();
      let loginUserIndex = ORFilterOptns.Provider.findIndex((user) => {
        return (
          user.text.toLowerCase() ==
          props.spcontext.pageContext.user.displayName.toLowerCase()
        );
      });
      let loginUserData = ORFilterOptns.Provider.splice(loginUserIndex, 1);

      ORFilterOptns.Provider.sort(sortFilterKeys);
      ORFilterOptns.Provider.unshift(loginUserData[0]);
      ORFilterOptns.Provider = usersOrderFunction(ORFilterOptns.Provider);
      ORFilterOptns.Provider.unshift({ key: "All", text: "All" });
    } else {
      ORFilterOptns.Provider.shift();
      ORFilterOptns.Provider.sort(sortFilterKeys);
      ORFilterOptns.Provider = usersOrderFunction(ORFilterOptns.Provider);
      ORFilterOptns.Provider.unshift({ key: "All", text: "All" });
    }

    setORFilterDrpDown(ORFilterOptns);
  };

  // getData function
  const getOrgReportData = (): void => {
    let _ORdata: IData[] = [];
    sharepointWeb.lists
      .getByTitle(OrgReportListName)
      .items.select("*", "FieldValuesAsText/DueDate")
      .expand("FieldValuesAsText") 
      .orderBy("ID", false)
      .top(5000)
      .get()
      .then((items: any[]) => {
 

        items.forEach((item: any) => {
          if (

            item.UserId === currentLoggedUserID   && (item.UserType.filter(x=>x=="Approver")).length>0 &&
            // ((item.Frequency == "Term" && item.TimePeriod == OR_Term) ||
            //   (item.Frequency == "Weekly" && item.TimePeriod == OR_Week) ||
            //   (item.Frequency == "Monthly" && item.TimePeriod == OR_Month)) &&
            item.Status == "Not actioned" 
            &&
            item.DocLink
          ) {
            let ProviderDetails = null;
            if (item.ProviderId) {
              ProviderDetails = allPeoples.filter((people) => {
                return people.ID == item.ProviderId;
              })[0];
            }
            _ORdata.push({
              ID: item.ID,
              BA: item.BA,
              Title: item.Title,
              Frequency: item.Frequency,
              DueDate: moment(
                item["FieldValuesAsText"].DueDate,
                DateListFormat
              ).format(DateListFormat),
              Provider: ProviderDetails,
              ConfigID: item.ConfigID,
              DocLink: item.DocLink,
              TLink: item.TLink,
              Year: item.Year,
              TimePeriod: item.TimePeriod,
              Status: item.Status,
              GroupID: item.GroupID,
              UserType: item.UserType,
              ParentID: item.ParentID,
              MasterData: item.MasterData,
            });
          }
        });

        setORFilterData([..._ORdata]);
        sortORFilterData = _ORdata;
        setORData([..._ORdata]);
        sortORData = _ORdata;
        setORMasterData([..._ORdata]);
        reloadFilterDropdowns([..._ORdata]);
        paginateFunction(1, _ORdata);
        setORLoader("noLoader");
      })
      .catch((error) => {
        ORErrorFunction(error, "getOrgReportData");
      });
  };

  // column-sorting function
  const _onColumnClick = (
    ev: React.MouseEvent<HTMLElement>,
    column: IColumn
  ): void => {
    const tempORColumns = _ORApprovalRequestsColumn;
    const newColumns: IColumn[] = tempORColumns.slice();
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

    const newORData = _copyAndSort(
      sortORData,
      currColumn.fieldName!,
      currColumn.isSortedDescending
    );
    const newORFilterData = _copyAndSort(
      sortORFilterData,
      currColumn.fieldName!,
      currColumn.isSortedDescending
    );
    setORData([...newORData]);
    setORFilterData([...newORFilterData]);
    paginateFunction(1, [...newORFilterData]);
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
  // Function-Declaration Ends

  useEffect(() => {
    setORLoader("StartLoader");
    getOrgReportData();
  }, [ORReRender]);
  return (
    <div style={{ marginTop: "5px" }}>
      {ORLoader == "StartLoader" ? (
        // <div style={{ height: 100, width: 100, margin: "auto" }}>
        <CustomLoader />
      ) : (
        // </div>
        <div>
          {/* Header-Section Starts */}
          <div
            style={{
              display: "flex",
              justifyContent: "space-between",
              alignItems: "end",
              flexWrap: "wrap",
            }}
          >
            <div
              style={{
                display: "flex",
                alignItems: "center",
                marginBottom: 10,
                flexWrap: "wrap",
              }}
            >
              <div>
                <Label styles={ORfilterLabelStyles}>Business area</Label>
                <Dropdown
                  placeholder="Select an option"
                  options={ORFilterDrpDown.BA}
                  selectedKey={ORFilter.BA}
                  styles={
                    ORFilter.BA == "All"
                      ? ORDropdownStyles
                      : ORActiveDropdownStyles
                  }
                  onChange={(e, option: any) => {
                    onChangeFilter("BA", option["key"]);
                  }}
                />
              </div>
              <div>
                <Label styles={ORfilterLabelStyles}>Title</Label>
                <Dropdown
                  placeholder="Select an option"
                  options={ORFilterDrpDown.Title}
                  selectedKey={ORFilter.Title}
                  styles={
                    ORFilter.Title == "All"
                      ? ORDropdownStyles
                      : ORActiveDropdownStyles
                  }
                  onChange={(e, option: any) => {
                    onChangeFilter("Title", option["key"]);
                  }}
                />
              </div>
              <div>
                <Label styles={ORfilterLabelStyles}>Frequency</Label>
                <Dropdown
                  placeholder="Select an option"
                  options={ORFilterDrpDown.Frequency}
                  selectedKey={ORFilter.Frequency}
                  styles={
                    ORFilter.Frequency == "All"
                      ? ORDropdownStyles
                      : ORActiveDropdownStyles
                  }
                  onChange={(e, option: any) => {
                    onChangeFilter("Frequency", option["key"]);
                  }}
                />
              </div>
              <div>
                <Label styles={ORfilterLabelStyles}>Provider</Label>
                <Dropdown
                  placeholder="Select an option"
                  options={ORFilterDrpDown.Provider}
                  selectedKey={ORFilter.Provider}
                  styles={
                    ORFilter.Provider == "All"
                      ? ORDropdownStyles
                      : ORActiveDropdownStyles
                  }
                  onChange={(e, option: any) => {
                    onChangeFilter("Provider", option["key"]);
                  }}
                />
              </div>
              <div>
                <Icon
                  iconName="Refresh"
                  title="Click to reset"
                  className={ORiconStyleClass.refresh}
                  onClick={() => {
                    setORFilterData([...ORMasterData]);
                    setORData([...ORMasterData]);
                    sortORFilterData = ORMasterData;
                    sortORData = ORMasterData;
                    paginateFunction(1, ORMasterData);
                    setORFilter({ ...ORFilterKeys });
                    setORColumns(_ORApprovalRequestsColumn);
                  }}
                />
              </div>
            </div>
            <div
              style={{
                display: "flex",
                alignItems: "center",
                justifyContent: "left",
                // paddingTop: 16,
                paddingBottom: "10px",
              }}
            >
              <Label
                onClick={() => {
                  generateExcel();
                }}
                style={{
                  backgroundColor: "#EBEBEB",
                  padding: "7px 15px",
                  cursor: "pointer",
                  fontSize: "12px",
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                  borderRadius: "3px",
                  color: "#1D6F42",
                  marginRight: 10,
                }}
              >
                <Icon
                  style={{
                    color: "#1D6F42",
                  }}
                  iconName="ExcelDocument"
                  className={ORiconStyleClass.export}
                />
                Export as XLS
              </Label>
            </div>
          </div>

          {/* Header-Section Ends */}
          {/* Body-Section Starts */}
          <div>
            {/* DetailList-Section Starts */}
            <DetailsList
              items={ORDisplayData}
              columns={ORColumns}
              styles={{
                root: {
                  ".ms-DetailsHeader-cellTitle": {
                    // justifyContent: "center !important",
                  },
                  ".ms-DetailsRow-cell": {
                    display: "flex !important",
                    alignItems: "center !important",
                  },
                },
              }}
              setKey="set"
              layoutMode={DetailsListLayoutMode.justified}
              selectionMode={SelectionMode.none}
            />
            {/* DetailList-Section Ends */}
          </div>
          {ORFilterData.length > 0 ? (
            <div
              style={{
                display: "flex",
                justifyContent: "center",
                margin: "10px 0",
              }}
            >
              <Pagination
                currentPage={ORCurrentPage}
                totalPages={
                  ORFilterData.length > 0
                    ? Math.ceil(ORFilterData.length / totalPageItems)
                    : 1
                }
                onChange={(page) => {
                  paginateFunction(page, ORFilterData);
                }}
              />
            </div>
          ) : (
            <div
              style={{
                display: "flex",
                justifyContent: "center",
                marginTop: "15px",
              }}
            >
              <Label style={{ color: "#2392B2" }}>No data found !!!</Label>
            </div>
          )}
          {/* Body-Section Ends */}
          {/* Modal-Section Starts */}
          <div>
            {ORPopup.condition ? (
              <Modal isOpen={ORPopup.condition} isBlocking={false}>
                <div style={{ padding: "15px 40px" }}>
                  <div
                    style={{
                      fontSize: 24,
                      textAlign: "center",
                      color: "#2392B2",
                      fontWeight: "600",
                      marginBottom: "20px",
                    }}
                  >
                    Approval
                  </div>
                  <div style={{ marginBottom: 10 }}>
                    <div>
                      <Label styles={ORfilterLabelStyles}>Status</Label>
                      <Dropdown
                        placeholder="Select an option"
                        options={
                          ORPopup.item.UserType.some(
                            (type) => type == "Approver"
                          )
                            ? ORModalDrpDwnOptns.ApproverStatus
                            : ORModalDrpDwnOptns.AudienceStatus
                        }
                        selectedKey={ORPopup.Status}
                        styles={
                          ORPopup.statusValidation
                            ? ORModalErrorDropdownStyles
                            : ORModalDropdownStyles
                        }
                        onChange={(e, option: any) => {
                          ORDataHandler("Status", option["key"]);
                        }}
                      />
                    </div>
                    <div>
                      <Label styles={ORfilterLabelStyles}>Comments</Label>
                      <TextField
                        styles={
                          ORPopup.commentsValidation
                            ? ORModalErrorTxtFieldStyles
                            : ORModalTxtFieldStyles
                        }
                        multiline={true}
                        resizable={false}
                        onChange={(e, value: string) => {
                          ORDataHandler("comments", value);
                        }}
                      />
                    </div>
                  </div>
                  <div className={styles.ORFileUploadModalBoxButtonSection}>
                    {ORPopup.overAllValidation ? (
                      <Label style={{ color: "#f00", fontWeight: 600 }}>
                        * All fields are mandatory
                      </Label>
                    ) : null}
                    <button
                      className={styles.ORModalBoxSubmitBtn}
                      onClick={(_) => {
                        if (ORLoader == "noLoader") {
                          setORLoader("onModalSubmit");
                          ORValidationFunction();
                        }
                      }}
                      style={{ display: "flex" }}
                    >
                      {ORLoader == "onModalSubmit" ? (
                        <Spinner />
                      ) : (
                        <span>
                          <Icon
                            iconName="Save"
                            style={{ position: "relative", top: 3, left: -8 }}
                          />
                          {"Submit"}
                        </span>
                      )}
                    </button>
                    <button
                      className={styles.ORModalBoxBackBtn}
                      onClick={(_) => {
                        setORPopup({
                          condition: false,
                          item: null,
                          Status: null,
                          comments: null,
                          statusValidation: false,
                          commentsValidation: false,
                          overAllValidation: false,
                        });
                      }}
                    >
                      <span>
                        <Icon
                          iconName="Close"
                          style={{ position: "relative", top: 3, left: -8 }}
                        />
                        Close
                      </span>
                    </button>
                  </div>
                </div>
              </Modal>
            ) : null}
          </div>
          {/* Modal-Section Ends */}
        </div>
      )}
    </div>
  );
};

export default OrgApprovalRequests;
