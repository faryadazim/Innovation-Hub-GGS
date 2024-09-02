import * as React from "react";
import { useState, useEffect } from "react";
import { IWeb, Web } from "@pnp/sp/webs";
import * as moment from "moment";
import { mergeStyles, mergeStyleSets } from "@fluentui/react/lib/Styling";
import {
  DetailsList,
  DetailsListLayoutMode,
  SelectionMode,
  Icon,
  Label,
  Dropdown,
  IDropdownStyles,
  TooltipHost,
  TooltipDelay,
  TooltipOverflowMode,
  DirectionalHint,
  IColumn,
  ILabelStyles,
  Persona,
  PersonaPresence,
  PersonaSize,
  Toggle,
  Stack,
  IStackTokens,
  IDropdownOption,
  SearchBox,
  ISearchBoxStyles,
} from "@fluentui/react";

import Service from "../components/Services";

import CustomLoader from "../../innovationHubIntranet/components/CustomLoader";
import "../ExternalRef/styleSheets/Styles.css";
import styles from "./WeeklyReport.module.scss";
import * as Excel from "exceljs/dist/exceljs.min.js";
import * as FileSaver from "file-saver";

import alertify from "alertifyjs";
import "alertifyjs/build/css/alertify.css";
let DateListFormat = "DD/MM/YYYY";

interface IPRDrpdwn {
  WeekOptns: IDropdownOption[];
  MonthOptns: IDropdownOption[];
  // POOptns: IDropdownOption[];
  YearOptns: IDropdownOption[];
  TermOptns: IDropdownOption[];
  TypeOfProjectOptns: IDropdownOption[];
  ProductOptns: IDropdownOption[];
  ProjectOptns: IDropdownOption[];
}
interface IData {
  ID: string;
  Term: any[];
  POID: number;
  PO: string;
  POEmail: string;
  Product: string;
  Project: string;
  TypeOfProject: string;
  Status: string;
  ShowAll: boolean;
  Developer: string;
  DeveloperEmail: string;
  StartDate: string;
  EndDate: string;
}

interface IFilter {
  showAll: any;
  Developer: string;
  Year: string;
  Term: string;
  Week: string;
  Month: string;
  TypeOfProject: string;
  Product: string;
  Project: string;
}

let globalMasterUserListData = [];
const WRProjectReport = (props) => {
  const sharepointWeb: IWeb = Web(props.URL);
  const ListName = props.ListName;
  let currentBA = props.BA;
  let PR_Week: number = moment().isoWeek();
  let PR_Year: number = moment().year();

  let loggeduseremail: string = props.spcontext.pageContext.user.email;

  const PRColumns: IColumn[] = [
    {
      key: "column1",
      name: "Developer",
      fieldName: "D",
      minWidth: 200,
      maxWidth: 230,
      onRender: (item) => (
        <div
          style={{
            display: "flex",
            marginTop: -5,
          }}
        >
          <div style={{ cursor: "pointer" }}>
            <Persona
              title={item.Developer}
              size={PersonaSize.size24}
              presence={PersonaPresence.none}
              imageUrl={
                "/_layouts/15/userphoto.aspx?size=S&username=" +
                `${item.DeveloperEmail}`
              }
            />
          </div>
          <Label style={{ marginTop: -3 }}>{item.Developer}</Label>
        </div>
      ),
    },
    {
      key: "column2",
      name: "Name of deliverable",
      fieldName: "Project",
      minWidth: 100,
      maxWidth: 150,

      onRender: (item) => (
        <>
          <TooltipHost
            id={item.ID}
            content={item.Project}
            overflowMode={TooltipOverflowMode.Parent}
          >
            <span aria-describedby={item.ID}>{item.Project}</span>
          </TooltipHost>
        </>
      ),
    },
    {
      key: "column3",
      name: "Product or solution",
      fieldName: "Product",
      minWidth: 100,
      maxWidth: 150,

      onRender: (item) => (
        <>
          <TooltipHost
            id={item.ID}
            content={item.Product}
            overflowMode={TooltipOverflowMode.Parent}
          >
            <span aria-describedby={item.ID}>{item.Product}</span>
          </TooltipHost>
        </>
      ),
    },
    {
      key: "Column4",
      name: "Start date",
      fieldName: "StartDate",
      minWidth: 80,
      maxWidth: 80,

      onRender: (item: IData) => <div>{item.StartDate}</div>,
    },
    {
      key: "Column5",
      name: "End date",
      fieldName: "EndDate",
      minWidth: 80,
      maxWidth: 80,

      onRender: (item: IData) => <div>{item.EndDate}</div>,
    },
    {
      key: "column6",
      name: "TOD",
      fieldName: "TypeOfProject",
      minWidth: 50,
      maxWidth: 60,
    },

    {
      key: "column7",
      name: "Term",
      fieldName: "Term",
      minWidth: 50,
      maxWidth: 60,

      onRender: (item) => <>{item.Term.join(",")}</>,
    },
    {
      key: "column8",
      name: "Client",
      fieldName: "PO",
      minWidth: 200,
      maxWidth: 230,

      onRender: (item: IData) => (
        <div
          style={{
            display: "flex",
            marginTop: -2,
          }}
        >
          <div style={{ cursor: "pointer" }}>
            <Persona
              title={item.PO}
              size={PersonaSize.size24}
              presence={PersonaPresence.none}
              imageUrl={
                "/_layouts/15/userphoto.aspx?size=S&username=" +
                `${item.POEmail}`
              }
            />
          </div>
          <Label style={{ marginTop: -3 }}>{item.PO}</Label>
        </div>
      ),
    },

    {
      key: "column9",
      name: "Status",
      fieldName: "Status",
      minWidth: 130,
      maxWidth: 150,

      onRender: (item) => (
        <>
          {item.Status == "Completed" ? (
            <div className={apStatusStyleClass.completed}>{item.Status}</div>
          ) : item.Status == "Scheduled" ? (
            <div className={apStatusStyleClass.scheduled}>{item.Status}</div>
          ) : item.Status == "On schedule" ? (
            <div className={apStatusStyleClass.onSchedule}>{item.Status}</div>
          ) : item.Status == "Behind schedule" ? (
            <div className={apStatusStyleClass.behindScheduled}>
              {item.Status}
            </div>
          ) : item.Status == "On hold" ? (
            <div className={apStatusStyleClass.Onhold}>{item.Status}</div>
          ) : (
            ""
          )}
        </>
      ),
    },
  ];

  const PRDrpDwnOptns: IPRDrpdwn = {
    // POOptns: [{ key: "All", text: "All" }],
    WeekOptns: [{ key: "All", text: "All" }],
    MonthOptns: [{ key: "All", text: "All" }],
    TermOptns: [{ key: "All", text: "All" }],
    TypeOfProjectOptns: [{ key: "All", text: "All" }],
    ProductOptns: [{ key: "All", text: "All" }],
    ProjectOptns: [{ key: "All", text: "All" }],
    YearOptns: [{ key: PR_Year.toString(), text: PR_Year.toString() }],
  };
  const PRFilterKeys: IFilter = {
    Developer: "",
    Year: PR_Year.toString(),
    showAll: false,
    Term: "All",
    Week: "All",
    Month: "All",
    TypeOfProject: "All",
    Product: "All",
    Project: "All",
  };

  const stackTokens: IStackTokens = { childrenGap: 10 };
  const PRiconStyle = mergeStyles({
    fontSize: 17,
    height: 14,
    width: 17,
    cursor: "pointer",
  });
  const PRiconStyleClass = mergeStyleSets({
    link: [{ color: "blue", margin: "0 0" }, PRiconStyle],
    delete: [{ color: "red", margin: "0 7px" }, PRiconStyle],
    edit: [{ color: "blue", margin: "0 7px" }, PRiconStyle],
    refresh: [
      {
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
    ],
    export: {
      color: "#038387",
      fontSize: "18px",
      height: 20,
      width: 20,
      cursor: "pointer",
      marginRight: 5,
    },
  });
  const buttonStyles = {
    root: {
      //background: 'transparent',
      //border: 'none',
      minWidth: "30px",
      padding: 0,
      marginRight: "10px",
    },
  };
  const apStatusStyle = mergeStyles({
    textAlign: "center",
    borderRadius: "25px",
    width: 125,
  });
  const apStatusStyleClass = mergeStyleSets({
    completed: [
      {
        fontWeight: "600",
        padding: "3px",
        color: "#438700",
        backgroundColor: "#D9FFB3",
      },
      apStatusStyle,
    ],
    scheduled: [
      {
        fontWeight: "600",
        padding: "3px",
        color: "#9f6700",
        backgroundColor: "#FFDB99",
      },
      apStatusStyle,
    ],
    onSchedule: [
      {
        fontWeight: "600",
        padding: "3px",
        color: "#B3B300 ",
        backgroundColor: "#FFFFB3",
      },
      apStatusStyle,
    ],
    behindScheduled: [
      {
        fontWeight: "600",
        padding: "3px",
        color: "#FF0000",
        backgroundColor: "#FFB3B3",
      },
      apStatusStyle,
    ],
    Onhold: [
      {
        fontWeight: "600",
        padding: "3px",
        color: "#773030",
        backgroundColor: "#e6b1b1",
      },
      apStatusStyle,
    ],
  });
  const DBShortDropdownStyles: Partial<IDropdownStyles> = {
    root: {
      width: 75,
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
    callout: {
      maxHeight: 300,
    },
  };
  const DBActiveShortDropdownStyles: Partial<IDropdownStyles> = {
    root: {
      width: 75,
      marginRight: "15px",
      backgroundColor: "#F5F5F7",
    },
    title: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
      color: "#038387",
      border: "2px solid #038387",
      borderRadius: 4,
      fontWeight: 600,
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
    callout: {
      maxHeight: 300,
    },
  };
  const DBDropdownStyles: Partial<IDropdownStyles> = {
    root: {
      width: 150,
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
    callout: {
      maxHeight: 300,
    },
  };
  const DBActiveDropdownStyles: Partial<IDropdownStyles> = {
    root: {
      width: 150,
      marginRight: "15px",
      backgroundColor: "#F5F5F7",
    },
    title: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
      color: "#038387",
      border: "2px solid #038387",
      borderRadius: 4,
      fontWeight: 600,
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
    callout: {
      maxHeight: 300,
    },
  };
  const PRSearchBoxStyles: Partial<ISearchBoxStyles> = {
    root: {
      width: 186,
      marginRight: 15,
      backgroundColor: "#F5F5F7",
      outline: "none",
      border: "1px solid #E8E8EA",
      borderRadius: 4,
    },
    icon: { fontSize: 12, color: "#000" },
  };
  const PRActiveSearchBoxStyles: Partial<ISearchBoxStyles> = {
    root: {
      width: 186,
      marginRight: 15,
      backgroundColor: "#F5F5F7",
      outline: "none",
      border: "2px solid #038387",
      borderRadius: 4,
    },
    field: { fontWeight: 600, color: "#038387" },
    icon: { fontSize: 12, color: "#038387" },
  };

  const [PRitems, setPRitems] = useState<IData[]>([]);
  const [group, setgroup] = useState([]);
  const [PRFilters, setPRFilters] = useState<IData[]>([]);
  const [PRColumn, SetPRColumn] = useState<IColumn[]>(PRColumns);
  const [PRFilterDropDown, setPRFilterDroppDown] =
    useState<IPRDrpdwn>(PRDrpDwnOptns);
  const [PRLoader, setPRLoader] = useState("noLoader");
  const [PRFilter, setPRFilter] = useState<IFilter>(PRFilterKeys);
  const getMasterUserListData = (year: number, _filterKeys: IFilter): void => {
    setPRLoader("StartLoader");

    sharepointWeb.lists
      .getByTitle("Master User List")
      .items.select("*,User/ID,User/EMail,User/Title")
      .expand("User")
      .filter(`BusinessArea eq '${currentBA}'`)
      .top(5000)
      .get()
      .then((items) => {
        globalMasterUserListData = [];

        items = items.filter((user) => {
          return user.UserId;
        });

        items.forEach((user) => {
          globalMasterUserListData.push({
            userID: user.User.ID,
            userName: user.User.Title,
            userEmail: user.User.EMail,
            userBA: user.BusinessArea,
          });
        });

        getPRData(year, _filterKeys);
      })
      .catch((err) => {
        PRErrorFunction(err, "MasterUserListData-getData");
      });
  };
  const getPRData = (year: number, filterKeys: IFilter) => {
    // setPRLoader('StartLoader')
    let _Sldata: IData[] = [];
    sharepointWeb.lists
      .getByTitle(ListName)
      .items.select(
        "*",
        "ProjectOwner/Title",
        "ProjectOwner/Id",
        "ProjectOwner/EMail",
        "ProjectLead/Title",
        "ProjectLead/Id",
        "ProjectLead/EMail",
        "Master_x0020_Project/Title",
        "Master_x0020_Project/Id",
        "FieldValuesAsText/StartDate",
        "FieldValuesAsText/PlannedEndDate"
      )
      .expand(
        "ProjectOwner",
        "ProjectLead",
        "Master_x0020_Project",
        "FieldValuesAsText"
      )
      .filter(`Year eq ${year}`)
      .top(5000)
      .get()
      .then((items) => {
        globalMasterUserListData.forEach((gblUser) => {
          items.forEach((item) => {
            let arrTerm = [];
            arrTerm.push(`${item.Term}`);

            if (item.ProjectLeadId && item.ProjectLead.length != 0) {
              if (
                item.ProjectLead.some((dev) => dev.EMail == gblUser.userEmail)
              ) {
                _Sldata.push({
                  ID: item.ID,
                  Term:
                    item.TermNew != null && item.TermNew.length > 0
                      ? [...item.TermNew]
                      : item.Term
                      ? [...arrTerm]
                      : [],
                  POID: item.ProjectOwnerId ? item.ProjectOwnerId : null,
                  PO: item.ProjectOwnerId ? item.ProjectOwner.Title : "",
                  POEmail: item.ProjectOwnerId ? item.ProjectOwner.EMail : "",
                  TypeOfProject: item.ProjectType,
                  Project: item.Title,
                  Product: item.Master_x0020_ProjectId
                    ? item.Master_x0020_Project.Title
                    : "",
                  Status: item.Status,
                  ShowAll:
                    item.Status == "Behind schedule" || item.Status == "On hold"
                      ? true
                      : false,
                  Developer: gblUser.userName,
                  DeveloperEmail: gblUser.userEmail,
                  StartDate: item.StartDate
                    ? moment(
                        item["FieldValuesAsText"].StartDate,
                        DateListFormat
                      ).format(DateListFormat)
                    : "",
                  EndDate: item.PlannedEndDate
                    ? moment(
                        item["FieldValuesAsText"].PlannedEndDate,
                        DateListFormat
                      ).format(DateListFormat)
                    : "",
                });
              }
            }
          });
        });

        console.log(_Sldata, "_Sldata");

        reloadFilterDropdowns([..._Sldata], true, "");
        filterFunction(_Sldata, filterKeys, "");
        setPRFilter({ ...filterKeys });

        // getFilterOptns()

        setPRitems(_Sldata);

        setPRLoader("noLoader");
      });
  };
  const groups = (records) => {
    let newRecords = [];
    records.forEach((rd, index) => {
      newRecords.push({
        Title: rd.Developer,
        indexValue: index,
      });
    });

    let varGroup = [];
    let Uniquelessons = newRecords.reduce(function (item, e1) {
      var matches = item.filter(function (e2) {
        return e1.Title === e2.Title;
      });

      if (matches.length == 0) {
        item.push(e1);
      }
      return item;
    }, []);

    Uniquelessons.forEach((ul) => {
      let lessonLength = newRecords.filter((arr) => {
        return arr.Title == ul.Title;
      }).length;
      varGroup.push({
        key: ul.Title,
        name: ul.Title,
        startIndex: ul.indexValue,
        count: lessonLength,
      });
    });
    console.log([...varGroup]);
    setgroup([...varGroup]);
  };
  const reloadFilterDropdowns = (data: IData[], onload, month): void => {
    data.forEach((item) => {
      if (
        PRDrpDwnOptns.TypeOfProjectOptns.findIndex((arr) => {
          return arr.key == item.TypeOfProject;
        }) == -1 &&
        item.TypeOfProject
      ) {
        PRDrpDwnOptns.TypeOfProjectOptns.push({
          key: item.TypeOfProject,
          text: item.TypeOfProject,
        });
      }
      if (
        PRDrpDwnOptns.ProjectOptns.findIndex((arr) => {
          return arr.key == item.Project;
        }) == -1 &&
        item.Project
      ) {
        PRDrpDwnOptns.ProjectOptns.push({
          key: item.Project,
          text: item.Project,
        });
      }
      if (
        PRDrpDwnOptns.ProductOptns.findIndex((arr) => {
          return arr.key == item.Product;
        }) == -1 &&
        item.Product
      ) {
        PRDrpDwnOptns.ProductOptns.push({
          key: item.Product,
          text: item.Product,
        });
      }
    });

    for (let j = 2020; j <= PR_Year; j++) {
      PRDrpDwnOptns.YearOptns.push({
        key: j.toString(),
        text: j.toString(),
      });
    }
    ["1", "2", "3", "4"].forEach((_item) => {
      if (
        PRDrpDwnOptns.TermOptns.findIndex((termOptn) => {
          return termOptn.key == _item;
        }) == -1 &&
        _item
      ) {
        PRDrpDwnOptns.TermOptns.push({
          key: _item,
          text: _item,
        });
      }
    });
    for (let j = 0; j < 12; j++) {
      let monthName = moment().month(j).format("MMMM");
      PRDrpDwnOptns.MonthOptns.push({
        key: j,
        text: monthName,
      });
    }
    let minweek;
    let maxWeek;
    if (onload) {
      minweek = 1;
      maxWeek = parseInt(PRFilter.Year) == PR_Year ? PR_Week : 53;
    } else {
      let startOfMonth = moment().month(month).startOf("month")["_d"];
      let endOfMonth = moment().month(month).endOf("month")["_d"];

      minweek = moment(startOfMonth).week();
      maxWeek = moment(endOfMonth).week();
      console.log(startOfMonth, endOfMonth, minweek, maxWeek);
    }
    for (let i = minweek; i <= maxWeek; i++) {
      PRDrpDwnOptns.WeekOptns.push({
        key: i.toString(),
        text: i.toString(),
      });
    }

    PRDrpDwnOptns.YearOptns.shift();

    console.log(PRDrpDwnOptns.YearOptns);
    // SetPRColumn(PRColumns);
    setPRFilterDroppDown({ ...PRDrpDwnOptns });
  };
  const onChangePRFilter = (
    key: string,
    option: string | boolean,
    onload: boolean
  ) => {
    let tempFilterKeys = PRFilter;

    tempFilterKeys[key] = option;
    setPRFilter({ ...tempFilterKeys });
    if (key == "Year") {
      getMasterUserListData(parseInt(tempFilterKeys.Year), tempFilterKeys);
    } else {
      filterFunction(PRitems, tempFilterKeys, key);
    }
  };
  const filterFunction = (data, filterKeys: IFilter, key) => {
    let _tempData: IData[] = data;
    let tempData =
      filterKeys.showAll == true
        ? _tempData
        : _tempData.filter((_data) => {
            return (
              _data.Status == "Behind schedule" || _data.Status == "On hold"
            );
          });
    if (filterKeys.TypeOfProject != "All") {
      tempData = tempData.filter((arr) => {
        return arr.TypeOfProject == filterKeys.TypeOfProject;
      });
    }
    if (filterKeys.Product != "All") {
      tempData = tempData.filter((arr) => {
        return arr.Product == filterKeys.Product;
      });
    }
    if (filterKeys.Project != "All") {
      tempData = tempData.filter((arr) => {
        return arr.Project == filterKeys.Project;
      });
    }
    if (filterKeys.Developer) {
      tempData = tempData.filter((arr) => {
        return arr.Developer.toLowerCase().includes(
          filterKeys.Developer.toLowerCase()
        );
      });
    }
    if (filterKeys.Term != "All") {
      let termArr = [];
      tempData.forEach((arr) => {
        if (arr.Term.length != 0) {
          if (arr.Term.some((term) => term == filterKeys.Term)) {
            termArr.push(arr);
          }
        }
      });
      tempData = [...termArr];
    }
    if (filterKeys.Month != "All") {
      tempData = tempData.filter((arr) => {
        let minMonth: number = moment(arr.StartDate, DateListFormat).month();
        let maxMonth: number = moment(arr.EndDate, DateListFormat).month();
        if (minMonth > maxMonth) {
          let firstDateofMonth = `${filterKeys.Month + 1}-01-${
            filterKeys.Year
          }`;
          minMonth = moment(firstDateofMonth).month();
        }
        let month: number = parseInt(filterKeys.Month);
        return month >= minMonth && month <= maxMonth;
      });
    }
    if (filterKeys.Week != "All") {
      tempData = tempData.filter((arr) => {
        let minWeek: number = moment(arr.StartDate, DateListFormat).week();
        let maxWeek: number = moment(arr.EndDate, DateListFormat).week();
        if (minWeek > maxWeek) {
          let firstDateofMonth = `${
            moment().isoWeek(Number(filterKeys.Week)).month() + 1
          }-01-${filterKeys.Year}`;
          minWeek = moment(firstDateofMonth).week();
        }
        let Week: number = parseInt(filterKeys.Week);
        return Week >= minWeek && Week <= maxWeek;
      });
    }

    console.log(tempData, "new");
    setPRFilters([...tempData]);
    groups([...tempData]);
    key == "Month" && filterKeys.Month != "All"
      ? reloadFilterDropdowns([], false, filterKeys.Month)
      : "";
    key == "Month" && filterKeys.Month == "All"
      ? reloadFilterDropdowns([], true, filterKeys.Month)
      : "";
  };

  const generateExcel = (): void => {
    let arrExport = PRFilters;
    const workbook = new Excel.Workbook();
    const worksheet = workbook.addWorksheet("My Sheet");
    worksheet.columns = [
      { header: "Developer", key: "D", width: 50 },
      { header: "Name of deliverable", key: "Project", width: 20 },
      { header: "Product or solution", key: "Product", width: 25 },
      { header: "Start date", key: "StartDate", width: 25 },
      { header: "End date", key: "EndDate", width: 25 },
      { header: "TOD", key: "TypeOfProject", width: 25 },
      { header: "Term", key: "Term", width: 25 },
      { header: "Client", key: "PO", width: 50 },
      { header: "Status", key: "Status", width: 25 },
    ];
    arrExport.forEach((item: IData) => {
      worksheet.addRow({
        Project: item.Project ? item.Project : "",
        Product: item.Product ? item.Product : "",
        StartDate: item.StartDate ? item.StartDate : "",
        EndDate: item.EndDate ? item.EndDate : "",
        TypeOfProject: item.TypeOfProject ? item.TypeOfProject : "",
        Term: item.Term ? item.Term : "",
        PO: item.PO ? item.PO : "",
        D: item.Developer,
        Status: item.Status ? item.Status : "",
      });
    });
    ["A1", "B1", "C1", "D1", "E1", "F1", "G1", "H1", "I1"].map((key) => {
      worksheet.getCell(key).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "00e8d1" },
      };
    });
    ["A1", "B1", "C1", "D1", "E1", "F1", "G1", "H1", "I1"].map((key) => {
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
          `Weeklyreport-${new Date().toLocaleString()}.xlsx`
        )
      )
      .catch((err) => console.log("Error writing excel export", err));
  };

  const ErrorPopup = () => (
    alertify.set("notifier", "position", "top-right"),
    alertify.error("Something when error, please contact system admin.")
  );

  const PRErrorFunction = (error: any, functionName: string): void => {
    console.log(error, functionName);

    let response = {
      ComponentName: "Weekly report - project report",
      FunctionName: functionName,
      ErrorMessage: JSON.stringify(error["message"]),
      Title: loggeduseremail,
    };

    Service.SPAddItem({ Listname: "Error Log", RequestJSON: response }).then(
      () => {
        setPRLoader("noLoader");
        ErrorPopup();
      }
    );
  };
  useEffect(() => {
    getMasterUserListData(PR_Year, PRFilterKeys);
  }, [currentBA]);

  return (
    <>
      {PRLoader == "StartLoader" ? (
        <CustomLoader />
      ) : (
        <div>
          <div
            style={{
              // display: "flex",
              // justifyContent: "space-between",
              marginTop: "10px",
              // flexWrap: "wrap",
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
                <Label>Developer</Label>
                <SearchBox
                  placeholder="Search developer"
                  styles={
                    PRFilter.Developer
                      ? PRActiveSearchBoxStyles
                      : PRSearchBoxStyles
                  }
                  value={PRFilter.Developer}
                  onChange={(e, value): void => {
                    onChangePRFilter("Developer", value, true);
                  }}
                />
              </div>
              <div>
                <Label>Deliverable</Label>
                <Dropdown
                  styles={
                    PRFilter.Project == "All"
                      ? DBDropdownStyles
                      : DBActiveDropdownStyles
                  }
                  dropdownWidth={"auto"}
                  options={PRFilterDropDown.ProjectOptns}
                  onChange={(e, option: any): void => {
                    onChangePRFilter("Project", option["key"], true);
                  }}
                  selectedKey={PRFilter.Project}
                />
              </div>
              <div>
                <Label>Product</Label>
                <Dropdown
                  styles={
                    PRFilter.Product == "All"
                      ? DBDropdownStyles
                      : DBActiveDropdownStyles
                  }
                  dropdownWidth={"auto"}
                  options={PRFilterDropDown.ProductOptns}
                  onChange={(e, option: any): void => {
                    onChangePRFilter("Product", option["key"], true);
                  }}
                  selectedKey={PRFilter.Product}
                />
              </div>
              <div>
                <Label>TOD</Label>
                <Dropdown
                  styles={
                    PRFilter.TypeOfProject == "All"
                      ? DBShortDropdownStyles
                      : DBActiveShortDropdownStyles
                  }
                  options={PRFilterDropDown.TypeOfProjectOptns}
                  onChange={(e, option: any): void => {
                    onChangePRFilter("TypeOfProject", option["key"], true);
                  }}
                  selectedKey={PRFilter.TypeOfProject}
                />
              </div>
              <div>
                <Label>Term</Label>
                <Dropdown
                  styles={
                    PRFilter.Term == "All"
                      ? DBShortDropdownStyles
                      : DBActiveShortDropdownStyles
                  }
                  options={PRFilterDropDown.TermOptns}
                  onChange={(e, option: any): void => {
                    onChangePRFilter("Term", option["key"], true);
                  }}
                  selectedKey={PRFilter.Term}
                />
              </div>
              <div>
                <Label>Year</Label>
                <Dropdown
                  styles={DBActiveShortDropdownStyles}
                  options={PRFilterDropDown.YearOptns}
                  onChange={(e, option: any): void => {
                    onChangePRFilter("Year", option["key"], true);
                  }}
                  selectedKey={PRFilter.Year}
                />
              </div>
              <div>
                <Label>Month</Label>
                <Dropdown
                  styles={
                    PRFilter.Month == "All"
                      ? DBDropdownStyles
                      : DBActiveDropdownStyles
                  }
                  options={PRFilterDropDown.MonthOptns}
                  onChange={(e, option: any): void => {
                    onChangePRFilter("Month", option["key"], true);
                  }}
                  selectedKey={PRFilter.Month}
                />
              </div>
              <div>
                <Label>Week</Label>
                <Dropdown
                  styles={
                    PRFilter.Week == "All"
                      ? DBShortDropdownStyles
                      : DBActiveShortDropdownStyles
                  }
                  options={PRFilterDropDown.WeekOptns}
                  onChange={(e, option: any): void => {
                    onChangePRFilter("Week", option["key"], true);
                  }}
                  selectedKey={PRFilter.Week}
                />
              </div>
              <div style={{ marginLeft: "10px", marginRight: "10px" }}>
                <Stack tokens={stackTokens}>
                  <Toggle
                    label="Show All"
                    styles={buttonStyles}
                    checked={PRFilter.showAll}
                    onChange={(e) => {
                      onChangePRFilter("showAll", !PRFilter.showAll, true);
                    }}
                  />
                </Stack>
              </div>
              <div>
                <Icon
                  iconName="Refresh"
                  title="Click to reset"
                  className={PRiconStyleClass.refresh}
                  onClick={() => {
                    SetPRColumn(PRColumns);
                    getMasterUserListData(PR_Year, PRFilterKeys);
                  }}
                />
              </div>
            </div>
            <div
              style={{
                display: "flex",
                alignItems: "center",
                justifyContent: "flex-end",
                paddingTop: 16,
              }}
            >
              <Label
                style={{
                  color: "#323130",
                  fontSize: "13px",
                  marginLeft: "10px",
                  fontWeight: "500",
                  marginRight: 10,
                }}
              >
                Number of records:{" "}
                <b style={{ color: "#038387" }}>{PRFilters.length}</b>
              </Label>
              <Label
                onClick={() => {
                  generateExcel();
                }}
                style={{
                  backgroundColor: "#EBEBEB",
                  padding: "7px 15px",
                  cursor: "pointer",
                  fontSize: 12,
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                  borderRadius: 3,
                  color: "#1D6F42",
                  // marginRight: 10,
                }}
              >
                <Icon
                  style={{
                    color: "#1D6F42",
                  }}
                  iconName="ExcelDocument"
                  className={PRiconStyleClass.export}
                />
                Export as XLS
              </Label>
            </div>
          </div>
          <div style={{ marginTop: "15px" }}>
            <DetailsList
              items={PRFilters}
              columns={PRColumn}
              groups={group}
              groupProps={{
                showEmptyGroups: true,
              }}
              styles={{
                root: {
                  ".ms-DetailsRow-cell": {
                    height: 40,
                  },
                  ".ms-DetailsHeader-cellTitle": {
                    background: "#03828711 !important",
                    color: "#038387 !important",
                  },
                },
              }}
              selectionMode={SelectionMode.none}
              data-is-scrollable={true}
              onShouldVirtualize={() => {
                return false;
              }}
              setKey="set"
              layoutMode={DetailsListLayoutMode.justified}
              onRenderRow={(data, defaultRender) => (
                <div>
                  {defaultRender({
                    ...data,

                    styles: {
                      root: {
                        background:
                          PRFilter.showAll && data.item.ShowAll == true
                            ? "#FFF2F2"
                            : "#fff",

                        selectors: {
                          "&:hover": {
                            background:
                              PRFilter.showAll && data.item.ShowAll == true
                                ? "#f5e3e3"
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
          {PRFilters.length > 0 ? null : (
            <div
              style={{
                display: "flex",
                justifyContent: "center",
                marginTop: "15px",
              }}
            >
              <Label style={{ color: "#2392B2", fontWeight: 600 }}>
                No data found !!!
              </Label>
            </div>
          )}
        </div>
      )}
    </>
  );
};

export default WRProjectReport;
