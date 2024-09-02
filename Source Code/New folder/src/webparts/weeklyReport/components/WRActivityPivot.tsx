import * as React from "react";
import { useState, useEffect } from "react";
import { IWeb, Web } from "@pnp/sp/webs";
import * as moment from "moment";
import {
  DetailsList,
  DetailsListLayoutMode,
  Dropdown,
  IColumn,
  Icon,
  IDropdownOption,
  IDropdownStyles,
  ILabelStyles,
  ISearchBoxStyles,
  IStackTokens,
  Label,
  mergeStyles,
  mergeStyleSets,
  Persona,
  PersonaPresence,
  PersonaSize,
  SearchBox,
  SelectionMode,
  Stack,
  Toggle,
} from "@fluentui/react";

import Service from "../components/Services";

import alertify from "alertifyjs";
import "alertifyjs/build/css/alertify.css";

import { IPeoplelist } from "../../innovationHubIntranet/components/IInnovationHubIntranetProps";
import Pagination from "office-ui-fabric-react-pagination";
import CustomLoader from "../../innovationHubIntranet/components/CustomLoader";

import * as Excel from "exceljs/dist/exceljs.min.js";
import * as FileSaver from "file-saver";

interface IProps {
  context: any;
  spcontext: any;
  graphContent: any;
  URL: string;
  peopleList: IPeoplelist[];
  BA: string;
}
interface IData {
  ID: string;
  UserID: number;
  UserName: string;
  UserEmail: string;
  Product: string;
  Project: string;
  TaskType: string;
  Task: string;
  PlannedHours: string;
  ActualHours: string;
  ShowAll: boolean;
}
interface ISolDrpdwn {
  UserOptns: string;
  ProductOptns: IDropdownOption[];
  WeekOptns: IDropdownOption[];
  YearOptns: IDropdownOption[];
  ProjectOptns: IDropdownOption[];
}
interface IwrapFilterKeys {
  user: string;
  product: string;
  week: string;
  year: string;
  project: string;
  showAll: boolean;
}

let columnSortArr: IData[] = [];
let userData = [];
let PBData = [];
let APBData = [];

let sortMasterData: IData[] = [];
let sortData: IData[] = [];

let sortMasterDataDL = [];
let sortDataDL = [];

let gblFilterKeys;

const WRActivityPivot = (props: IProps): JSX.Element => {
  const sharepointWeb: IWeb = Web(props.URL);
  let loggeduseremail = props.spcontext.pageContext.user.email;
  const wrapAllitems: IData[] = [];
  const allPeoples: any[] = props.peopleList;
  const currentBA = props.BA;

  let WRAP_Year: number = moment().year();
  let WRAP_Week: number = moment().isoWeek();
  let CurrentPage: number = 1;
  let totalPageItems: number = 10;

  const wrapColumns: IColumn[] = [
    {
      key: "column1",
      name: "User",
      fieldName: "UserName",
      minWidth: 250,
      maxWidth: 400,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
      onRender: (item) => (
        <div style={{ display: "flex" }}>
          <div
            style={{
              marginTop: "-6px",
            }}
            title={item.UserName}
          >
            <Persona
              size={PersonaSize.size24}
              presence={PersonaPresence.none}
              imageUrl={
                "/_layouts/15/userphoto.aspx?size=S&username=" +
                `${item.UserEmail}`
              }
            />
          </div>
          <div>
            <span title={item.Provider} style={{ fontSize: "13px" }}>
              {item.UserName}
            </span>
          </div>
        </div>
      ),
    },
    {
      key: "column5",
      name: "Planned hours",
      fieldName: "PlannedHours",
      minWidth: 110,
      maxWidth: 400,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
      onRender: (item: IData) => (
        <div>
          {item.PlannedHours.toString().match(/\./g)
            ? parseFloat(item.PlannedHours).toFixed(2)
            : item.PlannedHours}
        </div>
      ),
    },
    {
      key: "column6",
      name: "Actual hours",
      fieldName: "ActualHours",
      minWidth: 110,
      maxWidth: 400,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
      onRender: (item: IData) => (
        <div>
          {item.ActualHours.toString().match(/\./g)
            ? parseFloat(item.ActualHours).toFixed(2)
            : item.ActualHours}
        </div>
      ),
    },
    {
      key: "column7",
      name: "Action",
      fieldName: "Action",
      minWidth: 110,
      maxWidth: 250,
      onRender: (item) => (
        <>
          <Icon
            iconName="PageArrowRight"
            title="Click to history"
            style={{
              fontSize: 25,
              height: 14,
              width: 17,
              color: "#038387",
              cursor: "pointer",
            }}
            onClick={() => {
              // Dropdown Values
              let tempFilterOptions = { ...WRAPDropDownOptions };
              item.DetailReport.forEach((arr) => {
                if (
                  tempFilterOptions.ProjectOptns.findIndex((prj) => {
                    return prj.key == arr.Project;
                  }) == -1 &&
                  arr.Project
                ) {
                  tempFilterOptions.ProjectOptns.push({
                    key: arr.Project,
                    text: arr.Project,
                  });
                }
                if (
                  tempFilterOptions.ProductOptns.findIndex((prd) => {
                    return prd.key == arr.Product;
                  }) == -1 &&
                  arr.Product
                ) {
                  tempFilterOptions.ProductOptns.push({
                    key: arr.Product,
                    text: arr.Product,
                  });
                }
              });

              setWRAPDropDownOptions(tempFilterOptions);

              // Filter reset
              console.log(gblFilterKeys, "filter");
              let tempFilterKeys: IwrapFilterKeys = { ...gblFilterKeys };
              tempFilterKeys["project"] = "All";
              tempFilterKeys["product"] = "All";
              setWrapFilter({ ...tempFilterKeys });
              gblFilterKeys = tempFilterKeys;

              // DetailList
              setWrapDetailData(item.DetailReport);
              sortDataDL = item.DetailReport;
              setWrapMasterDetailData(item.DetailReport);
              setWrapUnsortMasterDetailData(item.DetailReport);
              sortMasterDataDL = item.DetailReport;
              setWrapDetailColumn(wrapDetailColumn);
              setDetailView({ condition: true, data: item });
              console.log(item.DetailReport, "data");
            }}
          />
        </>
      ),
    },
  ];
  const wrapDetailColumn: IColumn[] = [
    {
      key: "column1",
      name: "Product",
      fieldName: "Product",
      minWidth: 150,
      maxWidth: 250,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClickforDL(ev, column);
      },
    },
    {
      key: "column2",
      name: "Project or task",
      fieldName: "Project",
      minWidth: 100,
      maxWidth: 250,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClickforDL(ev, column);
      },
    },
    {
      key: "column3",
      name: "Task",
      fieldName: "Task",
      minWidth: 150,
      maxWidth: 400,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClickforDL(ev, column);
      },
    },
    {
      key: "column4",
      name: "Planned hours",
      fieldName: "PlannedHours",
      minWidth: 110,
      maxWidth: 350,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClickforDL(ev, column);
      },
    },
    {
      key: "column5",
      name: "Actual hours",
      fieldName: "ActualHours",
      minWidth: 110,
      maxWidth: 350,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClickforDL(ev, column);
      },
    },
    {
      key: "column6",
      name: "Start date",
      fieldName: "StartDate",
      minWidth: 110,
      maxWidth: 350,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClickforDL(ev, column);
      },
      onRender: (item) => (item.StartDate ? item.StartDate : null),
    },
    {
      key: "column7",
      name: "End date",
      fieldName: "EndDate",
      minWidth: 110,
      maxWidth: 350,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClickforDL(ev, column);
      },
      onRender: (item) => (item.EndDate ? item.EndDate : null),
    },
    {
      key: "column8",
      name: "Status",
      fieldName: "Status",
      minWidth: 120,
      maxWidth: 350,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClickforDL(ev, column);
      },
      onRender: (item) => (
        <>
          {item.Status == "Completed" ? (
            <div className={apStatusStyleClass.completed}>{item.Status}</div>
          ) : item.Status == "Scheduled" ? (
            <div className={apStatusStyleClass.scheduled}>{item.Status}</div>
          ) : item.Status == "Behind schedule" ? (
            <div className={apStatusStyleClass.behindScheduled}>
              {item.Status}
            </div>
          ) : item.Status == "On hold" ? (
            <div className={apStatusStyleClass.Onhold}>{item.Status}</div>
          ) : (
            <div className={apStatusStyleClass.onSchedule}>{item.Status}</div>
          )}
        </>
      ),
    },
  ];
  const wrapFilterKey: IwrapFilterKeys = {
    user: "",
    product: "All",
    week: WRAP_Week.toString(),
    year: WRAP_Year.toString(),
    project: "All",
    showAll: false,
  };
  const wrapDrpDwnOptns: ISolDrpdwn = {
    UserOptns: "",
    ProductOptns: [{ key: "All", text: "All" }],
    WeekOptns: [],
    YearOptns: [],
    ProjectOptns: [{ key: "All", text: "All" }],
  };

  // Design - Section
  const apStatusStyle = mergeStyles({
    textAlign: "center",
    borderRadius: "25px",
    width: 120,
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

  const wrapDropdownStyles: Partial<IDropdownStyles> = {
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

  const wrapActiveDropdownStyles: Partial<IDropdownStyles> = {
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
  };
  const wrapActiveSearchBoxStyles: Partial<ISearchBoxStyles> = {
    root: {
      width: 165,
      marginRight: "15px",
      backgroundColor: "#F5F5F7",
      outline: "none",
      border: "2px solid #038387",
      borderRadius: "4px",
      //marginTop: '3px',
    },
    field: { fontWeight: 600, color: "#038387" },
    icon: { fontSize: 14, color: "#038387" },
  };
  const WrapShortDropdownStyles: Partial<IDropdownStyles> = {
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
  };
  const wrapActiveShortDropdownStyles: Partial<IDropdownStyles> = {
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
  const wrapSearchBoxStyles: Partial<ISearchBoxStyles> = {
    root: {
      width: 165,
      marginRight: "15px",
      backgroundColor: "#F5F5F7",
      outline: "none",
      border: "1px solid #E8E8EA",
      borderRadius: "4px",
      //marginTop: '3px',
    },
    icon: { fontSize: 14, color: "#000" },
  };
  const wrapLabelStyles: Partial<ILabelStyles> = {
    root: {
      width: 165,
      //marginTop: 5,
      marginRight: 10,
      fontSize: 13,
      color: "#323130",
    },
  };
  const stackTokens: IStackTokens = { childrenGap: 10 };
  const wrapiconStyle = mergeStyles({
    fontSize: 17,
    height: 14,
    width: 17,
    cursor: "pointer",
  });
  const wrapiconStyleClass = mergeStyleSets({
    link: [{ color: "blue", margin: "0 0" }, wrapiconStyle],
    delete: [{ color: "red", margin: "0 7px" }, wrapiconStyle],
    edit: [{ color: "blue", margin: "0 7px" }, wrapiconStyle],
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
    pblink: [
      {
        color: "white",
        fontSize: "18px",
        height: 22,
        width: 22,
        cursor: "pointer",
        backgroundColor: "#038387",
        padding: 5,
        marginLeft: 10,
        borderRadius: 2,
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
      },
    ],
    export: [
      {
        color: "black",
        fontSize: "18px",
        height: 20,
        width: 20,
        cursor: "pointer",
        marginRight: 5,
      },
    ],
  });
  const wrapButtonStyles = {
    root: {
      //background: 'transparent',
      //border: 'none',
      minWidth: "30px",
      padding: 0,
      marginRight: "10px",
    },
  };
  const wrapIconStyleClass = mergeStyleSets({
    historyBackIcon: {
      color: "#000",
      fontSize: "16px",
      height: 20,
      width: 20,
      cursor: "pointer",
      marginRight: 5,
      fontWeight: 600,
      marginTop: 8,
    },
  });

  const WraplabelStyles = mergeStyleSets({
    NORLabel: [
      {
        color: "#323130",
        fontSize: "13px",
        marginLeft: "10px",
        fontWeight: "500",
        marginRight: 10,
      },
    ],
  });
  // let currentpage: number = 1
  // let totalPageItems: number = 10

  //UseState-Section
  // const [WrapReRender, setWrapReRender] = useState<boolean>(false);

  const [WrapDispalyData, setWrapDispalyData] = useState<IData[]>([]);
  const [WrapData, setWrapData] = useState<IData[]>([]);
  const [WrapMasterData, setWrapMasterData] = useState<IData[]>([]);

  const [WrapDetailData, setWrapDetailData] = useState<IData[]>([]);
  const [WrapMasterDetailData, setWrapMasterDetailData] = useState<IData[]>([]);
  const [WrapUnsortMasterDetailData, setWrapUnsortMasterDetailData] = useState<
    IData[]
  >([]);

  const [WRAPDropDownOptions, setWRAPDropDownOptions] =
    useState<ISolDrpdwn>(wrapDrpDwnOptns);

  const [WrapFilter, setWrapFilter] = useState<IwrapFilterKeys>(wrapFilterKey);
  const [WrapLoader, setWrapLoader] = useState("noLoader");
  const [DetailView, setDetailView] = useState({ condition: false, data: {} });

  const [WrapColumns, setWrapColumns] = useState<IColumn[]>(wrapColumns);
  const [WrapCurrentPage, setWrapCurrentPage] = useState<number>(CurrentPage);

  const [WrapDetailColumn, setWrapDetailColumn] =
    useState<IColumn[]>(wrapDetailColumn);

  const _onColumnClick = (
    ev: React.MouseEvent<HTMLElement>,
    column: IColumn
  ): void => {
    const tempapColumns = wrapColumns;
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

    const newSlData = _copyAndSort(
      sortMasterData,
      currColumn.fieldName!,
      currColumn.isSortedDescending
    );
    const newSlFilterData = _copyAndSort(
      sortData,
      currColumn.fieldName!,
      currColumn.isSortedDescending
    );
    setWrapMasterData([...newSlData]);
    setWrapData([...newSlFilterData]);
    paginateFunction(1, [...newSlFilterData]);
  };
  const _onColumnClickforDL = (
    ev: React.MouseEvent<HTMLElement>,
    column: IColumn
  ): void => {
    const tempapColumns = wrapDetailColumn;
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

    const newSlData = _copyAndSort(
      sortMasterDataDL,
      currColumn.fieldName!,
      currColumn.isSortedDescending
    );
    const newSlFilterData = _copyAndSort(
      sortDataDL,
      currColumn.fieldName!,
      currColumn.isSortedDescending
    );
    setWrapMasterDetailData([...newSlData]);
    setWrapDetailData([...newSlFilterData]);
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

  const generateExcel = (): void => {
    let arrExport = WrapData;
    const workbook = new Excel.Workbook();
    const worksheet = workbook.addWorksheet("My Sheet");
    worksheet.columns = [
      { header: "User", key: "User", width: 25 },
      { header: "Planned Hours", key: "PlannedHours", width: 25 },
      { header: "Actual Hours", key: "ActualHours", width: 25 },
    ];
    arrExport.forEach((item: IData) => {
      worksheet.addRow({
        User: item.UserName ? item.UserName : "",
        PlannedHours: item.PlannedHours ? item.PlannedHours : 0,
        ActualHours: item.ActualHours ? item.ActualHours : 0,
      });
    });
    ["A1", "B1", "C1"].map((key) => {
      worksheet.getCell(key).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "00e8d1" },
      };
    });
    ["A1", "B1", "C1"].map((key) => {
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
  const getThresholdData = (
    listName: string,
    filterCondition: string,
    weekNumber: number,
    year: number,
    onload: boolean
  ): void => {
    sharepointWeb.lists
      .getByTitle(listName)
      .renderListDataAsStream({
        ViewXml: filterCondition,
      })
      .then((data) => {
        listName == "ProductionBoard"
          ? PBData.push(...data.Row)
          : listName == "ActivityProductionBoard"
          ? APBData.push(...data.Row)
          : null;

        if (data.NextHref) {
          getPagedValues(
            listName,
            filterCondition,
            data.NextHref,
            weekNumber,
            year,
            onload
          );
        } else {
          listName == "ProductionBoard"
            ? getActivityProductionBoard(weekNumber, year, onload)
            : listName == "ActivityProductionBoard"
            ? getActivityPivot(onload)
            : null;
        }
      })
      .catch((err: string) => {
        WrapErrorFunction(err, "getThresholdData");
      });
  };
  const getPagedValues = (
    listName: string,
    filterCondition: string,
    nextHref: string,
    weekNumber: number,
    year: number,
    onload: boolean
  ): void => {
    sharepointWeb.lists
      .getByTitle(listName)
      .renderListDataAsStream({
        ViewXml: filterCondition,
        Paging: nextHref.substring(1),
      })
      .then((data) => {
        listName == "ProductionBoard"
          ? PBData.push(...data.Row)
          : listName == "ActivityProductionBoard"
          ? APBData.push(...data.Row)
          : null;

        if (data.NextHref) {
          getPagedValues(
            listName,
            filterCondition,
            data.NextHref,
            weekNumber,
            year,
            onload
          );
        } else {
          listName == "ProductionBoard"
            ? getActivityProductionBoard(weekNumber, year, onload)
            : listName == "ActivityProductionBoard"
            ? getActivityPivot(onload)
            : null;
        }
      })
      .catch((err: string) => {
        WrapErrorFunction(err, "getPagedValues");
      });
  };
  const getWrapData = (week: number, year: number, onload: boolean): void => {
    setWrapLoader("StartLoader");
    userData = [];
    PBData = [];
    APBData = [];
    const sortFilterKeys = (a, b) => {
      if (a.Title < b.Title) {
        return -1;
      }
      if (a.Title > b.Title) {
        return 1;
      }
      return 0;
    };
    sharepointWeb.lists
      .getByTitle("Master User List")
      .items.select("*,User/EMail,User/Title")
      .expand("User")
      .filter(`BusinessArea eq '${currentBA}'`)
      .top(5000)
      .get()
      .then((items) => {
        items = items.filter((item) => {
          return item.UserId != null;
        });
        userData.push(...items);
        userData.sort(sortFilterKeys);
        getProdutionBoard(week, year, onload);
      })
      .catch((err: string) => {
        WrapErrorFunction(err, "getWrapData");
      });
  };
  const getProdutionBoard = (week: number, year: number, onload: boolean) => {
    let Filtercondition = `
    <View Scope='RecursiveAll'>
      <Query>
      
         <OrderBy>
           <FieldRef Name='ID' Ascending='FALSE'/>
         </OrderBy>
         <Where>
         <And>
         <Eq>
            <FieldRef Name='Week' />
            <Value Type='Number'>${week}</Value>
         </Eq>
         <Eq>
            <FieldRef Name='Year' />
            <Value Type='Number'>${year}</Value>
         </Eq>
         </And>
         </Where>
      </Query>
      <ViewFields>
      <FieldRef Name='AnnualPlanID_x003a_Title' />
      <FieldRef Name='Project' />
      <FieldRef Name='Product' />
      <FieldRef Name='PlannedHours' />
      <FieldRef Name='ActualHours' />
      <FieldRef Name='Developer' />
      <FieldRef Name='Status' />
      <FieldRef Name='Title' />
      <FieldRef Name='StartDate' />
      <FieldRef Name='EndDate' />
      </ViewFields>
      <RowLimit Paged='TRUE'>5000</RowLimit>
    </View>`;

    getThresholdData("ProductionBoard", Filtercondition, week, year, onload);
  };
  const getActivityProductionBoard = (
    week: number,
    year: number,
    onload: boolean
  ) => {
    let Filtercondition = `
    <View Scope='RecursiveAll'>
      <Query>
      
         <OrderBy>
           <FieldRef Name='ID' Ascending='FALSE'/>
         </OrderBy>
         <Where>
         <And>
         <Eq>
            <FieldRef Name='Week' />
            <Value Type='Number'>${week}</Value>
         </Eq>
         <Eq>
            <FieldRef Name='Year' />
            <Value Type='Number'>${year}</Value>
         </Eq>
         </And>
         </Where>
      </Query>
      <ViewFields>
      <FieldRef Name='Project' />
      <FieldRef Name='Product' />
      <FieldRef Name='PlannedHours' />
      <FieldRef Name='ActualHours' />
      <FieldRef Name='Developer' />
      <FieldRef Name='Status' />
      <FieldRef Name='Steps' />
      <FieldRef Name='StartDate' />
      <FieldRef Name='EndDate' />
      </ViewFields>
      <RowLimit Paged='TRUE'>5000</RowLimit>
    </View>`;

    getThresholdData(
      "ActivityProductionBoard",
      Filtercondition,
      week,
      year,
      onload
    );
  };
  const getActivityPivot = (onload: boolean): void => {
    let tempArr = [];
    userData.forEach((item) => {
      let PB_PH = HoursCalculation(item.User.EMail, "PlannedHours", PBData);
      let APB_PH = HoursCalculation(item.User.EMail, "PlannedHours", APBData);
      let PB_AH = HoursCalculation(item.User.EMail, "ActualHours", PBData);
      let APB_AH = HoursCalculation(item.User.EMail, "ActualHours", APBData);
      let details = item.UserId
        ? DetailReport(item.User.EMail, PBData, APBData)
        : null;
      tempArr.push({
        UserID: item.UserId,
        UserName: item.UserId ? item.User.Title : "",
        UserEmail: item.UserId ? item.User.EMail : "",
        PlannedHours: PB_PH + APB_PH,
        ActualHours: PB_AH + APB_AH,
        DetailReport: details,
        ShowAll: Math.abs(PB_PH + APB_PH - (PB_AH + APB_AH)) > 5 ? true : false,
      });
    });
    console.log(tempArr);
    setWrapMasterData(tempArr);
    sortMasterData = tempArr;

    if (onload) {
      // ShowAll filter
      let tempArrwithFilter = tempArr.filter((arr) => {
        // return arr.PlannedHours != arr.ActualHours;
        return Math.abs(arr.PlannedHours - arr.ActualHours) > 5;
      });
      setWrapData(tempArrwithFilter);
      sortData = tempArrwithFilter;
      paginateFunction(1, [...tempArrwithFilter]);

      // Filter Choices

      let maxWeek = parseInt(WrapFilter.year) == WRAP_Year ? WRAP_Week : 53;

      for (var i = 1; i <= maxWeek; i++) {
        wrapDrpDwnOptns.WeekOptns.push({
          key: i.toString(),
          text: i.toString(),
        });
      }
      for (var i = 2020; i <= WRAP_Year; i++) {
        wrapDrpDwnOptns.YearOptns.push({
          key: i.toString(),
          text: i.toString(),
        });
      }

      setWRAPDropDownOptions(wrapDrpDwnOptns);
      setWrapFilter({ ...wrapFilterKey });
      gblFilterKeys = wrapFilterKey;
    } else {
      // filter
      let tempFilterKeys: IwrapFilterKeys = { ...WrapFilter };
      if (!tempFilterKeys.showAll) {
        tempArr = tempArr.filter((arr) => {
          return arr.PlannedHours != arr.ActualHours;
        });
      }

      if (tempFilterKeys.user != "") {
        tempArr = tempArr.filter((arr) => {
          return arr.UserName.toLowerCase().includes(
            tempFilterKeys.user.toLowerCase()
          );
        });
      }
      setWrapData([...tempArr]);
      sortData = tempArr;
      paginateFunction(1, [...tempArr]);
    }

    setWrapLoader("noLoader");
  };
  const wrapFilterData = (key: string, option: any, mainScreen): void => {
    let tempFilterKeys: IwrapFilterKeys = { ...WrapFilter };
    tempFilterKeys[key] = option;
    setWrapFilter({ ...tempFilterKeys });
    gblFilterKeys = tempFilterKeys;
    if (mainScreen) {
      if (key == "week" || key == "year") {
        getWrapData(
          parseInt(tempFilterKeys.week),
          parseInt(tempFilterKeys.year),
          false
        );
      } else {
        let arrBeforeFilter: IData[] = WrapMasterData;
        if (!tempFilterKeys.showAll) {
          arrBeforeFilter = arrBeforeFilter.filter((arr) => {
            return arr.PlannedHours != arr.ActualHours;
          });
        }

        if (tempFilterKeys.user != "") {
          arrBeforeFilter = arrBeforeFilter.filter((arr) => {
            return arr.UserName.toLowerCase().includes(
              tempFilterKeys.user.toLowerCase()
            );
          });
        }
        setWrapData([...arrBeforeFilter]);
        sortData = arrBeforeFilter;
        paginateFunction(1, [...arrBeforeFilter]);
      }
    } else {
      let arrBeforeFilter: IData[] = WrapMasterDetailData;
      if (tempFilterKeys.project != "All") {
        arrBeforeFilter = arrBeforeFilter.filter((arr) => {
          return arr.Project == tempFilterKeys.project;
        });
      }

      if (tempFilterKeys.product != "All") {
        arrBeforeFilter = arrBeforeFilter.filter((arr) => {
          return arr.Product == tempFilterKeys.product;
        });
      }
      setWrapDetailData([...arrBeforeFilter]);
      sortDataDL = arrBeforeFilter;
    }
    console.log(WrapFilter);
  };
  const HoursCalculation = (userEmail, field, data) => {
    let sum = 0;
    data = data.filter((arr) => {
      return arr.Developer ? arr.Developer[0].email == userEmail : null;
    });
    data.forEach((arr) => {
      parseFloat(arr[field]) ? (sum += parseFloat(arr[field])) : 0;
    });
    return sum;
  };
  const DetailReport = (userEmail, pbdata, apbdata) => {
    let tempArr = [];
    pbdata = pbdata.filter((arr) => {
      return arr.Developer ? arr.Developer[0].email == userEmail : null;
    });
    apbdata = apbdata.filter((arr) => {
      return arr.Developer ? arr.Developer[0].email == userEmail : null;
    });

    pbdata.forEach((arr) => {
      tempArr.push({
        Product: arr.Product.length > 0 ? arr.Product[0].lookupValue : "",
        Project: arr.AnnualPlanID_x003a_Title
          ? arr.AnnualPlanID_x003a_Title
          : arr.Project,
        Task: arr.Title,
        PlannedHours: arr.PlannedHours,
        ActualHours: arr.ActualHours,
        StartDate: arr.StartDate,
        EndDate: arr.EndDate,
        Status: arr.Status != "" ? arr.Status : "Scheduled",
      });
    });

    apbdata.forEach((arr) => {
      tempArr.push({
        Product: arr.Product,
        Project: arr.Project,
        Task: arr.Steps,
        PlannedHours: arr.PlannedHours,
        ActualHours: arr.ActualHours,
        StartDate: arr.StartDate,
        EndDate: arr.EndDate,
        Status: arr.Status != "" ? arr.Status : "Scheduled",
      });
    });

    return tempArr;
  };

  const paginateFunction = (pagenumber, data): void => {
    if (data.length > 0) {
      let lastIndex: number = pagenumber * totalPageItems;
      let firstIndex: number = lastIndex - totalPageItems;
      let paginatedItems = data.slice(firstIndex, lastIndex);
      CurrentPage = pagenumber;
      setWrapDispalyData(paginatedItems);
      setWrapCurrentPage(pagenumber);
    } else {
      setWrapDispalyData([]);
      setWrapCurrentPage(1);
    }
  };

  const ErrorPopup = () => (
    alertify.set("notifier", "position", "top-right"),
    alertify.error("Something when error, please contact system admin.")
  );
  const WrapErrorFunction = (error: any, functionName: string): void => {
    console.log(error, functionName);

    let response = {
      ComponentName: "Weekly report - activity pivot",
      FunctionName: functionName,
      ErrorMessage: JSON.stringify(error["message"]),
      Title: loggeduseremail,
    };

    Service.SPAddItem({ Listname: "Error Log", RequestJSON: response }).then(
      () => {
        setWrapLoader("noLoader");
        ErrorPopup();
      }
    );
  };

  //Function-Section Ends
  useEffect(() => {
    getWrapData(WRAP_Week, WRAP_Year, true);
  }, [currentBA]);

  return (
    <>
      <div>
        {DetailView.condition ? (
          <div>
            <div style={{ padding: 10, marginTop: "20px" }}>
              <Icon
                iconName="ChromeBack"
                className={wrapIconStyleClass.historyBackIcon}
                styles={{
                  root: {
                    transform: "translateY(3px)",
                  },
                }}
                onClick={() => {
                  // wrapFilterData("showAll", false, true);
                  setDetailView({ condition: false, data: {} });
                }}
              />
              <label
                style={{
                  fontSize: "18px",
                  marginLeft: "10px",
                  fontWeight: 600,
                }}
              >
                Report of {DetailView.data["UserName"]}
              </label>
            </div>
            <div
              style={{
                display: "flex",
                alignItems: "center",
                marginTop: "15px",
                marginBottom: "10px",
              }}
            >
              <div style={{ marginRight: "25px" }}>
                <label
                  style={{
                    marginRight: "5px",
                    color: "#000",
                    fontSize: "14px",
                    fontWeight: "500",
                  }}
                >
                  Planned hours :
                </label>
                <label
                  htmlFor=""
                  style={{
                    color: "#2392b2",
                    fontSize: "14px",
                    fontWeight: "500",
                  }}
                >
                  {DetailView.data["PlannedHours"].toString().match(/\./g)
                    ? parseFloat(DetailView.data["PlannedHours"]).toFixed(2)
                    : DetailView.data["PlannedHours"]}
                </label>
              </div>
              <div style={{ marginRight: "25px" }}>
                <label
                  style={{
                    marginRight: "5px",
                    color: "#000",
                    fontSize: "14px",
                    fontWeight: "500",
                  }}
                >
                  Actual hours :
                </label>
                <label
                  htmlFor=""
                  style={{
                    color: "#2392b2",
                    fontSize: "14px",
                    fontWeight: "500",
                  }}
                >
                  {DetailView.data["ActualHours"].toString().match(/\./g)
                    ? parseFloat(DetailView.data["ActualHours"]).toFixed(2)
                    : DetailView.data["ActualHours"]}
                </label>
              </div>
              <div style={{ marginRight: "25px" }}>
                <label
                  style={{
                    marginRight: "5px",
                    color: "#000",
                    fontSize: "14px",
                    fontWeight: "500",
                  }}
                >
                  Week:
                </label>
                <label
                  htmlFor=""
                  style={{
                    color: "#2392b2",
                    fontSize: "14px",
                    fontWeight: "500",
                  }}
                >
                  {WrapFilter.week}
                </label>
              </div>
              <div>
                <label
                  style={{
                    marginRight: "5px",
                    color: "#000",
                    fontSize: "14px",
                    fontWeight: "500",
                  }}
                >
                  Year :
                </label>
                <label
                  htmlFor=""
                  style={{
                    color: "#2392b2",
                    fontSize: "14px",
                    fontWeight: "500",
                  }}
                >
                  {WrapFilter.year}
                </label>
              </div>
            </div>
            <div style={{ display: "flex" }}>
              <div>
                <Label>Project or task</Label>
                <Dropdown
                  placeholder="Select an option"
                  options={WRAPDropDownOptions.ProjectOptns}
                  selectedKey={WrapFilter.project}
                  styles={
                    WrapFilter.project != "All"
                      ? wrapActiveDropdownStyles
                      : wrapDropdownStyles
                  }
                  onChange={(e, option: any) => {
                    wrapFilterData("project", option["key"], false);
                  }}
                />
              </div>
              <div>
                <Label>Product</Label>
                <Dropdown
                  placeholder="Select an option"
                  options={WRAPDropDownOptions.ProductOptns}
                  selectedKey={WrapFilter.product}
                  styles={
                    WrapFilter.product != "All"
                      ? wrapActiveDropdownStyles
                      : wrapDropdownStyles
                  }
                  onChange={(e, option: any) => {
                    wrapFilterData("product", option["key"], false);
                  }}
                />
              </div>
              <div>
                <Icon
                  iconName="Refresh"
                  title="Click to reset"
                  className={wrapiconStyleClass.refresh}
                  onClick={() => {
                    let tempFilterKeys: IwrapFilterKeys = { ...WrapFilter };
                    tempFilterKeys["project"] = "All";
                    tempFilterKeys["product"] = "All";
                    setWrapMasterDetailData(WrapUnsortMasterDetailData);
                    setWrapFilter({ ...tempFilterKeys });
                    gblFilterKeys = tempFilterKeys;
                    setWrapDetailData([...WrapUnsortMasterDetailData]);
                    sortDataDL = WrapUnsortMasterDetailData;
                    setWrapDetailColumn(wrapDetailColumn);
                  }}
                />
              </div>
            </div>
            <div style={{ marginTop: "15px" }}>
              <DetailsList
                items={WrapDetailData}
                //compact={isCompactMode}
                columns={WrapDetailColumn}
                selectionMode={SelectionMode.none}
                //getKey={this._getKey}
                setKey="set"
                styles={
                  WrapDetailData.length > 0
                    ? {
                        root: {
                          ".ms-DetailsRow-cell": {
                            height: 40,
                          },
                          ".ms-DetailsHeader-cellTitle": {
                            background: "#03828711 !important",
                            color: "#038387 !important",
                          },
                        },
                      }
                    : null
                }
                layoutMode={DetailsListLayoutMode.justified}
              />
              {WrapDetailData.length > 0 ? null : (
                <div
                  style={{
                    display: "flex",
                    justifyContent: "center",
                    marginTop: "15px",
                  }}
                >
                  <Label style={{ color: "#2392B2", fontWeight: 600 }}>
                    No data Found !!!
                  </Label>
                </div>
              )}
            </div>
          </div>
        ) : WrapLoader == "StartLoader" ? (
          <CustomLoader />
        ) : (
          <div>
            <div
              style={{
                display: "flex",
                justifyContent: "space-between",
                marginTop: "10px",
                flexWrap: "wrap",
              }}
            >
              <div
                style={{
                  display: "flex",
                  paddingTop: "10px",
                  flexWrap: "wrap",
                }}
              >
                <div>
                  <Label styles={wrapLabelStyles}>User</Label>
                  <SearchBox
                    placeholder="Search user"
                    styles={
                      WrapFilter.user
                        ? wrapActiveSearchBoxStyles
                        : wrapSearchBoxStyles
                    }
                    value={WrapFilter.user}
                    onChange={(e, value) => {
                      wrapFilterData("user", value, true);
                    }}
                  />
                </div>

                <div>
                  <Label>Week</Label>
                  <Dropdown
                    placeholder="Select an option"
                    options={WRAPDropDownOptions.WeekOptns}
                    selectedKey={WrapFilter.week}
                    styles={
                      WrapFilter.week
                        ? wrapActiveShortDropdownStyles
                        : WrapShortDropdownStyles
                    }
                    onChange={(e, option: any): void => {
                      wrapFilterData("week", option["key"], true);
                    }}
                  />
                </div>
                <div>
                  <Label>Year</Label>
                  <Dropdown
                    placeholder="Select an option"
                    options={WRAPDropDownOptions.YearOptns}
                    selectedKey={WrapFilter.year}
                    styles={
                      WrapFilter.year
                        ? wrapActiveShortDropdownStyles
                        : WrapShortDropdownStyles
                    }
                    onChange={(e, option: any): void => {
                      wrapFilterData("year", option["key"], true);
                    }}
                  />
                </div>

                <div style={{ marginLeft: "10px", marginRight: "10px" }}>
                  <Stack tokens={stackTokens}>
                    <Toggle
                      label="Show all"
                      styles={wrapButtonStyles}
                      checked={WrapFilter.showAll}
                      //defaultChecked
                      //onText="On"
                      //offText="Off"
                      onChange={(e) => {
                        wrapFilterData("showAll", !WrapFilter.showAll, true);
                      }}
                    />
                  </Stack>
                </div>
                <div>
                  <Icon
                    iconName="Refresh"
                    title="Click to reset"
                    className={wrapiconStyleClass.refresh}
                    onClick={() => {
                      setWrapFilter({ ...wrapFilterKey });
                      gblFilterKeys = wrapFilterKey;
                      setWrapColumns(wrapColumns);
                      getWrapData(WRAP_Week, WRAP_Year, true);
                    }}
                  />
                </div>
              </div>
              <div
                style={{
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "left",
                  paddingTop: 16,
                }}
              >
                <Label className={WraplabelStyles.NORLabel}>
                  Number of records:{" "}
                  <b style={{ color: "#038387" }}>{WrapData.length}</b>
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
                    className={wrapiconStyleClass.export}
                  />
                  Export as XLS
                </Label>
              </div>
            </div>
            <div style={{ marginTop: "15px" }}>
              <DetailsList
                items={WrapDispalyData}
                columns={WrapColumns}
                selectionMode={SelectionMode.none}
                setKey="set"
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
                layoutMode={DetailsListLayoutMode.justified}
                onRenderRow={(data, defaultRender) => (
                  <div>
                    {defaultRender({
                      ...data,

                      styles: {
                        root: {
                          background:
                            WrapFilter.showAll && data.item.ShowAll == true
                              ? "#FFF2F2"
                              : "#fff",

                          selectors: {
                            "&:hover": {
                              background:
                                WrapFilter.showAll && data.item.ShowAll == true
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
            {WrapData.length > 0 ? (
              <div
                style={{
                  display: "flex",
                  justifyContent: "center",
                  margin: "10px 0",
                }}
              >
                <Pagination
                  currentPage={WrapCurrentPage}
                  totalPages={
                    WrapData.length > 0
                      ? Math.ceil(WrapData.length / totalPageItems)
                      : 1
                  }
                  onChange={(page) => {
                    paginateFunction(page, WrapData);
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
                <Label style={{ color: "#2392B2", fontWeight: 600 }}>
                  No data found !!!
                </Label>
              </div>
            )}
            <div
              style={{
                display: "flex",
                justifyContent: "center",
                margin: "10px 0",
              }}
            ></div>
          </div>
        )}
      </div>
    </>
  );
};

export default WRActivityPivot;
