import * as React from "react";
import { useState, useEffect } from "react";
import * as moment from "moment";
import { Web } from "@pnp/sp/webs";
import {
  DetailsList,
  DetailsListLayoutMode,
  SelectionMode,
  Icon,
  Label,
  Dropdown,
  IDropdownStyles,
  Persona,
  PersonaPresence,
  PersonaSize,
  SearchBox,
  ISearchBoxStyles,
  IDatePickerStyles,
  TooltipHost,
  TooltipOverflowMode,
  IColumn,
} from "@fluentui/react";

import Service from "../components/Services";

import "../ExternalRef/styleSheets/Styles.css";
import styles from "./InnovationHubIntranet.module.scss";
import Pagination from "office-ui-fabric-react-pagination";
import { arraysEqual, IDetailsListStyles } from "office-ui-fabric-react";
import { mergeStyles, mergeStyleSets } from "@fluentui/react/lib/Styling";
import CustomLoader from "./CustomLoader";
import alertify from "alertifyjs";
import "alertifyjs/build/css/alertify.css";

let sortPAPData = [];
let sortPAPFilter = [];

const ProductActivityPlan = (props: any) => {
  const sharepointWeb = Web(props.URL);
  const ActivityPlanID = props.ActivityPlanID;
  const allPeoples = props.peopleList;

  let loggeduseremail: string = props.spcontext.pageContext.user.email;

  let CurrentPage = 1;
  let totalPageItems = 10;

  const _PAPColumns = [
    {
      key: "Column1",
      name: "Section",
      fieldName: "Lessons",
      minWidth: 200,
      maxWidth: 500,
      onColumnClick: (ev, column) => {
        _onColumnClick(ev, column);
      },
      onRender: (item) => (
        <>
          <TooltipHost
            id={item.ID}
            content={item.Lessons}
            overflowMode={TooltipOverflowMode.Parent}
          >
            <span aria-describedby={item.ID}>{item.Lessons}</span>
          </TooltipHost>
        </>
      ),
    },
    {
      key: "Column3",
      name: "Start Date",
      fieldName: "StartDate",
      minWidth: 100,
      maxWidth: 400,
      onColumnClick: (ev, column) => {
        _onColumnClick(ev, column);
      },
      onRender: (item) => moment(item.StartDate).format("DD/MM/YYYY"),
    },
    {
      key: "Column4",
      name: "End Date",
      fieldName: "EndDate",
      minWidth: 100,
      maxWidth: 400,
      onColumnClick: (ev, column) => {
        _onColumnClick(ev, column);
      },
      onRender: (item) => moment(item.EndDate).format("DD/MM/YYYY"),
    },
    {
      key: "Column5",
      name: "Developer",
      fieldName: "Developer",
      minWidth: 250,
      maxWidth: 500,
      onColumnClick: (ev, column) => {
        _onColumnClick(ev, column);
      },
      onRender: (item) => (
        <div
          style={{
            display: "flex",
            justifyContent: "flex-start",
          }}
        >
          <div
            style={{
              marginTop: "-6px",
            }}
          >
            <Persona
              size={PersonaSize.size32}
              presence={PersonaPresence.none}
              imageUrl={
                "/_layouts/15/userphoto.aspx?size=S&username=" +
                `${item.DeveloperEmail}`
              }
            />
          </div>
          <div>
            <span style={{ fontSize: "13px" }}>{item.Developer}</span>
          </div>
        </div>
      ),
    },
    {
      key: "Column7",
      name: "Link",
      fieldName: "Link",
      minWidth: 50,
      maxWidth: 50,
      onRender: (item) => (
        <>
          <Icon
            iconName="Link12"
            className={PAPIconStyleClass.link}
            onClick={() => {
              props.handleclick(
                "ProductActivityDeliveryPlan",
                item.Id,
                item.LessonID
              );
            }}
          />
        </>
      ),
    },
  ];
  const paplabelStyles = mergeStyleSets({
    titleLabel: [
      {
        color: "#676767",
        fontSize: "14px",
        marginRight: "10px",
        fontWeight: "400",
      },
    ],
    labelValue: [
      {
        color: "#0882A5",
        fontSize: "14px",
        marginRight: "10px",
      },
    ],
    inputLabels: [
      {
        color: "#323130",
        fontSize: "13px",
      },
    ],
    ErrorLabel: [
      {
        marginTop: "25px",
        marginLeft: "10px",
        fontWeight: "500",
        color: "#D0342C",
        fontSize: "13px",
      },
    ],
    NORLabel: [
      {
        marginTop: "25px",
        marginLeft: "10px",
        fontWeight: "500",
        color: "#323130",
        fontSize: "13px",
      },
    ],
  });
  const papIconStyleClass = mergeStyleSets({
    navArrow: [
      {
        cursor: "pointer",
        color: "#2392b2",
        fontSize: 24,
        marginTop: "3px",
        marginRight: 12,
      },
    ],
    navArrowDisabled: [
      {
        cursor: "pointer",
        color: "#ababab",
        fontSize: 24,
        marginTop: "3px",
        marginRight: 12,
      },
    ],
    link: [
      {
        fontSize: 17,
        height: 16,
        width: 16,
        color: "#fff",
        backgroundColor: "#038387",
        cursor: "pointer",
        padding: 8,
        borderRadius: 3,
        marginLeft: 10,
        ":hover": {
          backgroundColor: "#025d60",
        },
      },
    ],
    linkDisabled: [
      {
        fontSize: 18,
        height: 16,
        width: 19,
        color: "#fff",
        backgroundColor: "#ababab",
        cursor: "not-allowed",
        padding: 8,
        borderRadius: 3,
        marginLeft: 10,
      },
    ],
    refresh: [
      {
        fontSize: 18,
        height: 16,
        width: 19,
        color: "#fff",
        backgroundColor: "#038387",
        cursor: "pointer",
        padding: 8,
        borderRadius: 3,
        marginTop: 40,
        ":hover": {
          backgroundColor: "#025d60",
        },
      },
    ],
    save: [
      {
        fontSize: "18px",
        color: "#fff",
        paddingRight: 10,
      },
    ],
    edit: [
      {
        fontSize: "18px",
        color: "#fff",
        paddingRight: 10,
      },
    ],
  });
  const PAPDatePickerStyles: Partial<IDatePickerStyles> = {
    root: {
      width: 186,
      marginRight: 15,
      // marginTop: 5,
      backgroundColor: "#F5F5F7",
      borderRadius: 4,
    },
    wrapper: {
      borderRadius: "4px",
      ".ms-TextField-fieldGroup": {
        border: "none",
      },
      ".ms-TextField-field": {
        borderRadius: "4px !important",
      },
      ".readOnlyPlaceholder-203": {
        color: "#7C7C7C !important",
      },
    },
    readOnlyTextField: {
      backgroundColor: "#F5F5F7 !important",
      fontSize: 12,
      border: "1px solid #E8E8EA !important",
      borderRadius: 4,
    },
    icon: {
      fontSize: 18,
      color: "#7C7C7C",
    },
  };
  const PAPActiveDatePickerStyles: Partial<IDatePickerStyles> = {
    root: {
      width: 186,
      marginRight: 15,
      // marginTop: 5,
      backgroundColor: "#F5F5F7",
      borderRadius: 4,
    },
    wrapper: {
      borderRadius: "4px",
      ".ms-TextField-fieldGroup": {
        border: "none",
      },
      ".ms-TextField-field": {
        borderRadius: "4px !important",
      },
      ".readOnlyPlaceholder-203": {
        color: "#038387 !important",
      },
    },
    readOnlyTextField: {
      backgroundColor: "#F5F5F7 !important",
      fontSize: 12,
      border: "2px solid #038387 !important",
      borderRadius: 4,
      color: "#038387",
      fontWeight: 600,
    },
    icon: {
      fontSize: 18,
      color: "#038387",
      fontWeight: 600,
    },
  };
  const PAPDrpDwnOptns = {
    Lesson: [{ key: "All", text: "All" }],
    Developer: [{ key: "All", text: "All" }],
  };
  const PAPFilterKeys = {
    Lesson: "All",
    Developer: "All",
    StartDate: null,
    EndDate: null,
  };
  const PAPstatusStyle = mergeStyles({
    textAlign: "center",
    borderRadius: "25px",
    height: 17,
  });
  const PAPstatusStyleClass = mergeStyleSets({
    completed: [
      {
        fontWeight: "600",
        color: "#438700",
        backgroundColor: "#D9FFB3",
      },
      PAPstatusStyle,
    ],
    scheduled: [
      {
        fontWeight: "600",
        color: "#9f6700",
        backgroundColor: "#FFDB99",
      },
      PAPstatusStyle,
    ],
    onSchedule: [
      {
        fontWeight: "600",
        color: "#B3B300 ",
        backgroundColor: "#FFFFB3",
      },
      PAPstatusStyle,
    ],
    behindScheduled: [
      {
        fontWeight: "600",
        color: "#FF0000",
        backgroundColor: "#FFB3B3",
      },
      PAPstatusStyle,
    ],
    default: [
      {
        fontWeight: "600",
        position: "relative",
        backgroundColor: "#edebe9",
      },
      PAPstatusStyle,
    ],
    percentageText: [
      {
        position: "absolute !important",
        left: "50%",
        top: "50%",
        transform: "translate(-50%,-50%)",
        color: "#555",
      },
    ],
  });

  // detailslist
  const gridStyles: Partial<IDetailsListStyles> = {
    root: {
      selectors: {
        "& [role=grid]": {
          display: "flex",
          flexDirection: "column",
          ".ms-DetailsRow-fields": {
            alignItems: "stretch",
            height: 38,
          },
        },
      },
    },
    headerWrapper: {
      flex: "0 0 auto",
    },
    contentWrapper: {
      flex: "1 1 auto",
      overflowY: "auto",
      overflowX: "hidden",
    },
  };
  const PAPlabelStyles = mergeStyleSets({
    titleLabel: [
      {
        color: "#676767",
        fontSize: "14px",
        marginRight: "10px",
        fontWeight: "400",
      },
    ],
    labelValue: [
      {
        color: "#0882A5",
        fontSize: "14px",
        marginRight: "10px",
      },
    ],
    inputLabels: [
      {
        color: "#323130",
        fontSize: "13px",
      },
    ],
    ErrorLabel: [
      {
        marginTop: "25px",
        marginLeft: "10px",
        fontWeight: "500",
        color: "#D0342C",
        fontSize: "13px",
      },
    ],
    NORLabel: [
      {
        marginTop: "25px",
        marginLeft: "10px",
        fontWeight: "500",
        color: "#323130",
        fontSize: "13px",
      },
    ],
  });
  const PAPdropdownStyles: Partial<IDropdownStyles> = {
    root: { width: 186, marginRight: 15 },
    title: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
      border: "1px solid #E8E8EA",
    },
    dropdownItemsWrapper: { backgroundColor: "#F5F5F7", fontSize: 12 },
    dropdownItemSelected: { backgroundColor: "#DCDCDC", fontSize: 12 },
    caretDown: { fontSize: 14, color: "#000" },
  };
  const PAPActivedropdownStyles: Partial<IDropdownStyles> = {
    root: { width: 186, marginRight: 15 },
    title: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
      color: "#038387",
      fontWeight: 600,
      border: "2px solid #038387",
    },
    dropdownItemsWrapper: { backgroundColor: "#F5F5F7", fontSize: 12 },
    dropdownItemSelected: { backgroundColor: "#DCDCDC", fontSize: 12 },
    caretDown: { fontSize: 14, color: "#038387", fontWeight: 600 },
  };
  const PAPSearchBoxStyles: Partial<ISearchBoxStyles> = {
    root: {
      width: 186,
      marginRight: "15px",
      backgroundColor: "#F5F5F7",
      outline: "none",
      border: "1px solid #E8E8EA",
      borderRadius: "4px",
      marginTop: "3px",
    },
    icon: { fontSize: 14, color: "#000" },
  };
  const PAPActiveSearchBoxStyles: Partial<ISearchBoxStyles> = {
    root: {
      width: 186,
      marginRight: "15px",
      backgroundColor: "#F5F5F7",
      outline: "none",
      border: "2px solid #038387",
      borderRadius: "4px",
      marginTop: "3px",
    },
    field: { fontWeight: 600, color: "#038387" },
    icon: { fontSize: 14, color: "#038387" },
  };
  const ATIconStyle = mergeStyles({
    fontSize: 17,
    height: 14,
    width: 17,
    cursor: "pointer",
  });
  const PAPIconStyleClass = mergeStyleSets({
    link: [{ color: "#2392B2", margin: "0" }, ATIconStyle],
    delete: [{ color: "#CB1E06", margin: "0 7px " }, ATIconStyle],
    edit: [{ color: "#2392B2", margin: "0 7px 0 0" }, ATIconStyle],
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

  const getActivityTemplate = (AllPeoples) => {
    let _PAPdata = [];
    sharepointWeb.lists
      .getByTitle("Activity Plan")
      .items.getById(ActivityPlanID)
      // .select("*", "Developer/Title", "Developer/Id", "Developer/EMail")
      // .expand("Developer")
      .get()
      .then(async (item) => {
        let lessons = item.Lessons.split(";");

        lessons.forEach((ls) => {
          _PAPdata.push({
            Id: item.ID,
            Title: item.Title,
            Project: item.Project,
            Product: item.Product,
            LessonID: ls.split("~")[0],
            Lessons: ls.split("~")[1],
            StartDate: ls.split("~")[2],
            EndDate: ls.split("~")[3],
            DeveloperId: ls.split("~")[4],
            Developer: AllPeoples.filter((ap) => {
              return ap.ID == ls.split("~")[4];
            })[0].text,
            DeveloperEmail: AllPeoples.filter((ap) => {
              return ap.ID == ls.split("~")[4];
            })[0].secondaryText,
          });
        });
        setPAPData([..._PAPdata]);
        sortPAPData = _PAPdata;
        sortPAPFilter = _PAPdata;
        setPAPMasterData([..._PAPdata]);
        paginateFunction(1, [..._PAPdata]);
        reloadFilterOptions([..._PAPdata]);
        setPAPLoader("noLoader");
      })
      .catch((err) => {
        PAPErrorFunction(err, "getActivityTemplate");
      });
  };
  const reloadFilterOptions = (data) => {
    let tempArrReload = data;

    const sortFilterKeys = (a, b) => {
      if (a.text < b.text) {
        return -1;
      }
      if (a.text > b.text) {
        return 1;
      }
      return 0;
    };

    tempArrReload.forEach((at) => {
      if (
        PAPDrpDwnOptns.Lesson.findIndex((cd) => {
          return cd.key == at.Lessons;
        }) == -1 &&
        at.Lessons
      ) {
        PAPDrpDwnOptns.Lesson.push({
          key: at.Lessons,
          text: at.Lessons,
        });
      }
      if (
        PAPDrpDwnOptns.Developer.findIndex((prd) => {
          return prd.key == at.Developer;
        }) == -1 &&
        at.Developer
      ) {
        PAPDrpDwnOptns.Developer.push({
          key: at.Developer,
          text: at.Developer,
        });
      }
    });

    if (
      PAPDrpDwnOptns.Developer.some((managerOptn) => {
        return (
          managerOptn.text.toLowerCase() ==
          props.spcontext.pageContext.user.displayName.toLowerCase()
        );
      })
    ) {
      PAPDrpDwnOptns.Developer.shift();
      let loginUserIndex = PAPDrpDwnOptns.Developer.findIndex((user) => {
        return (
          user.text.toLowerCase() ==
          props.spcontext.pageContext.user.displayName.toLowerCase()
        );
      });
      let loginUserData = PAPDrpDwnOptns.Developer.splice(loginUserIndex, 1);

      PAPDrpDwnOptns.Developer.sort(sortFilterKeys);
      PAPDrpDwnOptns.Developer.unshift(loginUserData[0]);
      PAPDrpDwnOptns.Developer.unshift({ key: "All", text: "All" });
    } else {
      PAPDrpDwnOptns.Developer.shift();
      PAPDrpDwnOptns.Developer.sort(sortFilterKeys);
      PAPDrpDwnOptns.Developer.unshift({ key: "All", text: "All" });
    }
    setPAPDropDownOptions(PAPDrpDwnOptns);
  };
  const PAPListFilter = (key, option) => {
    let tempArr = [...PAPData];
    let tempDpFilterKeys = { ...PAPFilterOptions };
    tempDpFilterKeys[`${key}`] = option;

    if (tempDpFilterKeys.Lesson != "All") {
      tempArr = tempArr.filter((arr) => {
        return arr.Lessons == tempDpFilterKeys.Lesson;
      });
    }
    if (tempDpFilterKeys.Developer != "All") {
      tempArr = tempArr.filter((arr) => {
        return arr.Developer == tempDpFilterKeys.Developer;
      });
    }
    if (tempDpFilterKeys.StartDate) {
      tempArr = tempArr.filter((arr) => {
        return (
          moment(arr.StartDate).format("DD/MM/YYYY") ==
          moment(tempDpFilterKeys.StartDate).format("DD/MM/YYYY")
        );
      });
    }

    if (tempDpFilterKeys.EndDate) {
      tempArr = tempArr.filter((arr) => {
        return (
          moment(arr.EndDate).format("DD/MM/YYYY") ==
          moment(tempDpFilterKeys.EndDate).format("DD/MM/YYYY")
        );
      });
    }
    sortPAPFilter = tempArr;
    paginateFunction(1, [...tempArr]);
    setPAPFilterOptions({ ...tempDpFilterKeys });
  };
  const paginateFunction = (pagenumber, data) => {
    if (data.length > 0) {
      let lastIndex: number = pagenumber * totalPageItems;
      let firstIndex: number = lastIndex - totalPageItems;
      let paginatedItems = data.slice(firstIndex, lastIndex);
      CurrentPage = pagenumber;
      setPAPDisplayData(paginatedItems);
      setPAPCurrentPage(pagenumber);
    } else {
      setPAPDisplayData([]);
      setPAPCurrentPage(1);
    }
  };
  const ErrorPopup = () => (
    alertify.set("notifier", "position", "top-right"),
    alertify.error("Something when error, please contact system admin.")
  );
  const PAPErrorFunction = (error: any, functionName: string) => {
    console.log(error);

    let response = {
      ComponentName: "Product activity plan",
      FunctionName: functionName,
      ErrorMessage: JSON.stringify(error["message"]),
      Title: loggeduseremail,
    };

    Service.SPAddItem({ Listname: "Error Log", RequestJSON: response }).then(
      () => {
        setPAPLoader("noLoader");
        ErrorPopup();
      }
    );
  };

  // Sorting Function
  const _onColumnClick = (
    ev: React.MouseEvent<HTMLElement>,
    column: IColumn
  ): void => {
    const tempapColumns = _PAPColumns;
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

    const newPAPData = _copyAndSort(
      sortPAPData,
      currColumn.fieldName!,
      currColumn.isSortedDescending
    );
    const newPAPFilter = _copyAndSort(
      sortPAPFilter,
      currColumn.fieldName!,
      currColumn.isSortedDescending
    );
    setPAPData([...newPAPData]);
    sortPAPFilter = newPAPFilter;
    paginateFunction(1, newPAPFilter);
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

  // Use State
  const [PAPReRender, setPAPReRender] = useState(false);
  const [PAPMasterData, setPAPMasterData] = useState([]);
  const [PAPData, setPAPData] = useState([]);
  const [PAPDisplayData, setPAPDisplayData] = useState([]);
  const [PAPcurrentPage, setPAPCurrentPage] = useState(CurrentPage);
  const [PAPDropDownOptions, setPAPDropDownOptions] = useState(PAPDrpDwnOptns);
  const [PAPFilterOptions, setPAPFilterOptions] = useState(PAPFilterKeys);
  const [PAPLoader, setPAPLoader] = useState("noLoader");
  const [PAPColumns, setPAPColumns] = useState(_PAPColumns);

  //Use Effect
  useEffect(() => {
    setPAPLoader("startUpLoader");
    getActivityTemplate(allPeoples);
  }, [PAPReRender]);

  return (
    <div style={{ padding: "5px 15px" }}>
      {PAPLoader == "startUpLoader" ? <CustomLoader /> : null}
      <div
        style={{
          display: "flex",
          alignItems: "flex-start",
          justifyContent: "space-between",
          marginBottom: 10,
          color: "#2392b2",
        }}
      >
        {/* Header Start */}
        <div className={styles.dpTitle}>
          <Icon
            iconName="NavigateBack"
            className={papIconStyleClass.navArrow}
            onClick={() => {
              props.handleclick("ProductActivityTemplate", null);
            }}
          />
          <Label style={{ color: "#000", fontSize: 24, padding: 0 }}>
            Product Activity Plan
          </Label>
        </div>
      </div>
      <div
        style={{
          display: "flex",
          alignItems: "center",
          justifyContent: "space-between",
          marginBottom: "10px",
        }}
      >
        <div className={styles.Section1}>
          <div
            style={{
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              marginRight: "15px",
            }}
          >
            <Label className={paplabelStyles.titleLabel}>Project :</Label>
            <Label className={paplabelStyles.labelValue}>
              {PAPData.length > 0 ? PAPData[0].Project : ""}
            </Label>
          </div>
          <div
            style={{
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              marginRight: "15px",
            }}
          >
            <Label className={paplabelStyles.titleLabel}>Product :</Label>
            <Label className={paplabelStyles.labelValue}>
              {PAPData.length > 0 ? PAPData[0].Product : ""}
            </Label>
          </div>
          <div
            style={{
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              marginRight: "15px",
            }}
          >
            <Label className={paplabelStyles.titleLabel}>Template :</Label>
            <Label className={paplabelStyles.labelValue}>
              {PAPData.length > 0 ? PAPData[0].Title : ""}
            </Label>
          </div>
        </div>
      </div>
      <div
        style={{
          display: "flex",
          alignItems: "center",
          justifyContent: "space-between",
          paddingBottom: "10px",
          flexWrap: "wrap",
        }}
      >
        <div className={styles.ddSection}>
          <div>
            <Label className={PAPlabelStyles.inputLabels}>Section</Label>
            <Dropdown
              selectedKey={PAPFilterOptions.Lesson}
              placeholder="Select an option"
              options={PAPDropDownOptions.Lesson}
              styles={
                PAPFilterOptions.Lesson != "All"
                  ? PAPActivedropdownStyles
                  : PAPdropdownStyles
              }
              onChange={(e, option: any) => {
                PAPListFilter("Lesson", option["key"]);
              }}
            />
          </div>
          {/* <div style={{ marginTop: 5 }}>
            <Label className={PAPlabelStyles.inputLabels}>Start date</Label>
            <DatePicker
              placeholder="Select a start date"
              formatDate={dateFormater}
              value={PAPFilterOptions.StartDate}
              styles={
                PAPFilterOptions.StartDate
                  ? PAPActiveDatePickerStyles
                  : PAPDatePickerStyles
              }
              onSelectDate={(value: any) => {
                PAPListFilter("StartDate", value);
              }}
            />
          </div>
          <div style={{ marginTop: 5 }}>
            <Label className={PAPlabelStyles.inputLabels}>End date</Label>
            <DatePicker
              placeholder="Select a end date"
              formatDate={dateFormater}
              value={PAPFilterOptions.EndDate}
              styles={
                PAPFilterOptions.EndDate
                  ? PAPActiveDatePickerStyles
                  : PAPDatePickerStyles
              }
              onSelectDate={(value: any) => {
                PAPListFilter("EndDate", value);
              }}
            />
          </div> */}
          <div>
            <Label className={PAPlabelStyles.inputLabels}>Developer</Label>
            <Dropdown
              selectedKey={PAPFilterOptions.Developer}
              placeholder="Select an option"
              options={PAPDropDownOptions.Developer}
              styles={
                PAPFilterOptions.Developer != "All"
                  ? PAPActivedropdownStyles
                  : PAPdropdownStyles
              }
              onChange={(e, option: any) => {
                PAPListFilter("Developer", option["key"]);
              }}
            />
          </div>
          <div>
            <Icon
              iconName="Refresh"
              title="Click to reset"
              className={PAPIconStyleClass.refresh}
              onClick={() => {
                setPAPFilterOptions({ ...PAPFilterKeys });
                setPAPData([...PAPMasterData]);
                sortPAPData = PAPMasterData;
                sortPAPFilter = PAPMasterData;
                paginateFunction(1, [...PAPMasterData]);
                setPAPColumns(_PAPColumns);
              }}
            />
          </div>
        </div>
        <div>
          <Label style={{ marginRight: 5 }}>
            Number of records :{" "}
            <span style={{ color: "#038387" }}>{PAPDisplayData.length}</span>
          </Label>
        </div>
      </div>
      <div>
        <DetailsList
          items={PAPDisplayData}
          columns={PAPColumns}
          setKey="set"
          layoutMode={DetailsListLayoutMode.justified}
          selectionMode={SelectionMode.none}
          // styles={gridStyles}
          styles={{ root: { width: "100%" } }}
        />
      </div>
      {PAPDisplayData.length > 0 ? (
        <div
          style={{
            display: "flex",
            justifyContent: "center",
            margin: "10px 0",
          }}
        >
          <Pagination
            currentPage={PAPcurrentPage}
            totalPages={
              PAPData.length > 0
                ? Math.ceil(PAPData.length / totalPageItems)
                : 1
            }
            onChange={(page) => {
              paginateFunction(page, PAPData);
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
          <Label style={{ color: "#2392B2" }}>No data Found !!!</Label>
        </div>
      )}
    </div>
  );
};
export default ProductActivityPlan;
