import * as React from "react";
import { useState, useEffect } from "react";
import * as moment from "moment";
import { Web } from "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import "@pnp/sp/site-users/web";
import {
  DetailsList,
  DetailsListLayoutMode,
  IDetailsListStyles,
  SelectionMode,
  Icon,
  Label,
  ILabelStyles,
  Dropdown,
  IDropdownStyles,
  Persona,
  PersonaPresence,
  PersonaSize,
  TooltipHost,
  TooltipOverflowMode,
  IColumn,
} from "@fluentui/react";

import Service from "../components/Services";

import alertify from "alertifyjs";
import "alertifyjs/build/css/alertify.css";

import "../ExternalRef/styleSheets/Styles.css";
import { FontIcon, IIconProps } from "@fluentui/react/lib/Icon";
import { mergeStyles, mergeStyleSets } from "@fluentui/react/lib/Styling";
import styles from "./InnovationHubIntranet.module.scss";
import CustomLoader from "./CustomLoader";

let columnSortArr = [];
let columnSortMasterArr = [];

const ProductActivityDeliveryPlan = (props: any) => {
  // Variable-Declaration-Section Starts
  const sharepointWeb = Web(props.URL);
  const activityPlan_ID = props.ActivityPlanID;
  const lesson_Id = props.pageType;
  const activityPlanListName = "Activity Plan";
  const adpListName = "Activity Delivery Plan";

  let loggeduseremail: string = props.spcontext.pageContext.user.email;

  const padpCurrentWeekNumber = moment().isoWeek();
  const padpCurrentYear = moment().year();

  const padpAllitems = [];
  const padpColumns = [
    {
      key: "Lesson",
      name: "Section",
      fieldName: "Lesson",
      minWidth: 250,
      maxWidth: 400,

      onRender: (item) => (
        <>
          <TooltipHost
            id={item.ID}
            content={item.Lesson}
            overflowMode={TooltipOverflowMode.Parent}
          >
            <span aria-describedby={item.ID}>{item.Lesson}</span>
          </TooltipHost>
        </>
      ),
    },
    {
      key: "Steps",
      name: "Steps",
      fieldName: "Steps",
      minWidth: 250,
      maxWidth: 300,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },

      onRender: (item) => (
        <>
          <TooltipHost
            id={item.ID}
            content={item.Steps}
            overflowMode={TooltipOverflowMode.Parent}
          >
            <span aria-describedby={item.ID}>{item.Steps}</span>
          </TooltipHost>
        </>
      ),
    },
    {
      key: "PH",
      name: "PH",
      fieldName: "PH",
      minWidth: 100,
      maxWidth: 250,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
    },
    {
      key: "Start",
      name: "Start",
      fieldName: "Start",
      minWidth: 150,
      maxWidth: 150,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },

      onRender: (item) => <>{moment(item.Start).format("DD/MM/YYYY")}</>,
    },
    {
      key: "End",
      name: "End",
      fieldName: "End",
      minWidth: 150,
      maxWidth: 150,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },

      onRender: (item) => <>{moment(item.End).format("DD/MM/YYYY")}</>,
    },
    {
      key: "Developer",
      name: "Developer",
      fieldName: "Developer",
      minWidth: 250,
      maxWidth: 350,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },

      onRender: (item) =>
        item.Developer ? (
          <div style={{ display: "flex" }}>
            <div
              style={{
                marginTop: "-6px",
              }}
              title={item.Developer ? item.Developer.name : ""}
            >
              <Persona
                size={PersonaSize.size32}
                presence={PersonaPresence.none}
                imageUrl={
                  "/_layouts/15/userphoto.aspx?size=S&username=" +
                  `${item.Developer.email}`
                }
              />
            </div>
            <div>
              <Label style={{ fontSize: "13px" }}>{item.Developer.name}</Label>
            </div>
          </div>
        ) : null,
    },
    {
      key: "Status",
      name: "Status",
      fieldName: "Status",
      minWidth: 150,
      maxWidth: 150,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },

      onRender: (item) => (
        <div>
          {item.Status == "Completed" ? (
            <div className={padpStatusStyles.completed}>{item.Status}</div>
          ) : item.Status == "Scheduled" ? (
            <div className={padpStatusStyles.scheduled}>{item.Status}</div>
          ) : item.Status == "On schedule" ? (
            <div className={padpStatusStyles.onSchedule}>{item.Status}</div>
          ) : item.Status == "Behind schedule" ? (
            <div className={padpStatusStyles.behindScheduled}>
              {item.Status}
            </div>
          ) : (
            ""
          )}
        </div>
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
  const padpDrpDwnOptns = {
    statusOptns: [{ key: "All", text: "All" }],
    developerOptns: [{ key: "All", text: "All" }],
    stepsOptns: [{ key: "All", text: "All" }],
    lessonOptns: [{ key: "All", text: "All" }],
  };
  const padpFilterKeys = {
    status: "All",
    developer: "All",
    step: "All",
    lesson: "All",
  };
  // Variable-Declaration-Section Ends
  // Styles-Section Starts
  const padpLabelStyles: Partial<ILabelStyles> = {
    root: {
      width: 165,
      marginTop: 5,
      marginRight: 10,
      fontSize: 13,
      color: "#323130",
    },
  };
  const padpDropdownStyles: Partial<IDropdownStyles> = {
    root: {
      width: 200,
      marginTop: 5,
      marginRight: 15,
      backgroundColor: "#F5F5F7",
      borderRadius: 4,
    },
    title: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
      color: "#7C7C7C",
      border: "1px solid #E8E8EA",
      borderRadius: 4,
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
  const padpActiveDropdownStyles: Partial<IDropdownStyles> = {
    root: {
      width: 200,
      marginTop: 5,
      marginRight: 15,
      backgroundColor: "#F5F5F7",
      borderRadius: 4,
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
    caretDown: { fontSize: 14, color: "#038387", fontWeight: 600 },
  };
  const gridStyles: Partial<IDetailsListStyles> = {
    root: {
      selectors: {
        "& [role=grid]": {
          display: "flex",
          flexDirection: "column",
          ".ms-DetailsRow-fields": {
            alignItems: "center",
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
      overflow: "hidden",
    },
  };
  const padpCommonStatusStyle = mergeStyles({
    textAlign: "center",
    borderRadius: 25,
    fontWeight: "600",
    padding: 3,
    width: 100,
    display: "flex",
    justifyContent: "center",
  });
  const padpStatusStyles = mergeStyleSets({
    completed: [
      {
        color: "#438700",
        backgroundColor: "#D9FFB3",
      },
      padpCommonStatusStyle,
    ],
    scheduled: [
      {
        color: "#9f6700",
        backgroundColor: "#FFDB99",
      },
      padpCommonStatusStyle,
    ],
    onSchedule: [
      {
        color: "#B3B300 ",
        backgroundColor: "#FFFFB3",
      },
      padpCommonStatusStyle,
    ],
    behindScheduled: [
      {
        color: "#FF0000",
        backgroundColor: "#FFB3B3",
      },
      padpCommonStatusStyle,
    ],
  });
  const padpIconStyleClass = mergeStyleSets({
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
  const padpCommonStyles = mergeStyleSets({
    titleLabel: {
      color: "#2392B2 !important",
      fontWeight: "500",
      fontSize: 17,
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
      fontSize: 13,
    },
    inputField: {
      margin: "10px 0",
    },
    dateGridValidationErrorLabel: {
      color: "#d0342c !important",
      fontWeight: 600,
      marginLeft: 20,
    },
  });
  // Styles-Section Ends
  // States-Declaration Starts
  const [padpReRender, setPadpReRender] = useState(true);
  const [activtyPlanItem, setActivtyPlanItem] = useState([]);
  const [padpUnsortMasterData, setPadpUnsortMasterData] =
    useState(padpAllitems);
  const [padpMasterData, setPadpMasterData] = useState(padpAllitems);
  const [padpData, setPadpData] = useState(padpAllitems);
  const [padpDropDownOptions, setPadpDropDownOptions] =
    useState(padpDrpDwnOptns);
  const [padpFilters, setPadpFilters] = useState(padpFilterKeys);
  const [padpLoader, setPadpLoader] = useState("noLoader");
  const [padpMasterColumns, setPadpMasterColumns] = useState(padpColumns);

  // States-Declaration Ends
  //Function-Section Starts
  const getActivityPlanItem = () => {
    let _padpItem = [];

    sharepointWeb.lists
      .getByTitle(activityPlanListName)
      .items.getById(activityPlan_ID)
      .get()
      .then((item) => {
        _padpItem.push({
          ID: item.Id ? item.Id : "",
          Lesson: item.Lessons ? item.Lessons : "",
          Project: item.Project ? item.Project : "",
          Product: item.Product ? item.Product : "",
          Types: item.Types ? item.Types : "",
        });

        setActivtyPlanItem([..._padpItem]);
        setPadpLoader("noLoader");
      })
      .catch((err) => {
        padpErrorFunction(err, "getActivityPlanItem");
      });
  };
  const padpGetData = () => {
    sharepointWeb.lists
      .getByTitle(adpListName)
      .items.filter(
        "ActivityPlanID eq '" +
          activityPlan_ID +
          "' and LessonID eq '" +
          lesson_Id +
          "' "
      )
      .select("*", "Developer/Title", "Developer/Id", "Developer/EMail")
      .expand("Developer")
      .orderBy("OrderId", true)
      .get()
      .then((items) => {
        if (items.length > 0) {
          items.forEach((item, index) => {
            padpAllitems.push({
              indexValue: index,
              ID: item.Id ? item.Id : "",
              Lesson: item.Lesson ? item.Lesson : "",
              Steps: item.Title ? item.Title : "",
              PH: item.PlannedHours ? item.PlannedHours : "",
              Project: item.Project ? item.Project : "",
              Start: item.StartDate ? item.StartDate : null,
              End: item.EndDate ? item.EndDate : null,
              Developer: item.DeveloperId
                ? {
                    name: item.Developer.Title,
                    id: item.Developer.Id,
                    email: item.Developer.EMail,
                  }
                : {
                    name: null,
                    id: null,
                    email: null,
                  },
              Status: item.Status ? item.Status : "",
              AH: item.ActualHours ? item.ActualHours : 0,
            });
          });

          padpGetAllOptions(padpAllitems);

          setPadpUnsortMasterData([...padpAllitems]);
          columnSortArr = padpAllitems;
          setPadpData([...columnSortArr]);
          columnSortMasterArr = padpAllitems;
          setPadpMasterData([...padpAllitems]);
          setPadpLoader("noLoader");
        }
      })
      .catch((err) => {
        padpErrorFunction(err, "padpGetData");
      });
  };
  const padpGetAllOptions = (allItems: any) => {
    allItems.forEach((item: any) => {
      if (
        padpDrpDwnOptns.statusOptns.findIndex((statusOptn) => {
          return statusOptn.key == item.Status;
        }) == -1 &&
        item.Status
      ) {
        padpDrpDwnOptns.statusOptns.push({
          key: item.Status,
          text: item.Status,
        });
      }

      if (
        padpDrpDwnOptns.developerOptns.findIndex((developerOptn) => {
          return developerOptn.key == item.Developer.name;
        }) == -1 &&
        item.Developer.name
      ) {
        padpDrpDwnOptns.developerOptns.push({
          key: item.Developer.name,
          text: item.Developer.name,
        });
      }

      if (
        padpDrpDwnOptns.stepsOptns.findIndex((stepsOptn) => {
          return stepsOptn.key == item.Steps;
        }) == -1 &&
        item.Steps
      ) {
        padpDrpDwnOptns.stepsOptns.push({
          key: item.Steps,
          text: item.Steps,
        });
      }

      if (
        padpDrpDwnOptns.lessonOptns.findIndex((lessonOptn) => {
          return lessonOptn.key == item.Lesson;
        }) == -1 &&
        item.Lesson
      ) {
        padpDrpDwnOptns.lessonOptns.push({
          key: item.Lesson,
          text: item.Lesson,
        });
      }
    });
    let unsortedFilterKeys = padpSortingFilterKeys(padpDrpDwnOptns);
    setPadpDropDownOptions({ ...unsortedFilterKeys });
  };
  const padpSortingFilterKeys = (unsortedFilterKeys: any) => {
    const sortFilterKeys = (a, b) => {
      if (a.text < b.text) {
        return -1;
      }
      if (a.text > b.text) {
        return 1;
      }
      return 0;
    };

    unsortedFilterKeys.statusOptns.shift();
    unsortedFilterKeys.statusOptns.sort(sortFilterKeys);
    unsortedFilterKeys.statusOptns.unshift({ key: "All", text: "All" });

    if (
      unsortedFilterKeys.developerOptns.some((managerOptn) => {
        return (
          managerOptn.text.toLowerCase() ==
          props.spcontext.pageContext.user.displayName.toLowerCase()
        );
      })
    ) {
      unsortedFilterKeys.developerOptns.shift();
      let loginUserIndex = unsortedFilterKeys.developerOptns.findIndex(
        (user) => {
          return (
            user.text.toLowerCase() ==
            props.spcontext.pageContext.user.displayName.toLowerCase()
          );
        }
      );
      let loginUserData = unsortedFilterKeys.developerOptns.splice(
        loginUserIndex,
        1
      );

      unsortedFilterKeys.developerOptns.sort(sortFilterKeys);
      unsortedFilterKeys.developerOptns.unshift(loginUserData[0]);
      unsortedFilterKeys.developerOptns.unshift({ key: "All", text: "All" });
    } else {
      unsortedFilterKeys.developerOptns.shift();
      unsortedFilterKeys.developerOptns.sort(sortFilterKeys);
      unsortedFilterKeys.developerOptns.unshift({ key: "All", text: "All" });
    }

    unsortedFilterKeys.stepsOptns.shift();
    unsortedFilterKeys.stepsOptns.sort(sortFilterKeys);
    unsortedFilterKeys.stepsOptns.unshift({ key: "All", text: "All" });

    unsortedFilterKeys.lessonOptns.shift();
    unsortedFilterKeys.lessonOptns.sort(sortFilterKeys);
    unsortedFilterKeys.lessonOptns.unshift({ key: "All", text: "All" });

    return unsortedFilterKeys;
  };
  const padpListFilter = (key: string, option: any) => {
    let arrBeforeFilter = [...padpMasterData];

    let tempFilterKeys = { ...padpFilters };
    tempFilterKeys[key] = option;

    if (tempFilterKeys.status != "All") {
      arrBeforeFilter = arrBeforeFilter.filter((arr) => {
        return arr.Status == tempFilterKeys.status;
      });
    }

    if (tempFilterKeys.developer != "All") {
      arrBeforeFilter = arrBeforeFilter.filter((arr) => {
        return arr.Developer.name == tempFilterKeys.developer;
      });
    }

    if (tempFilterKeys.step != "All") {
      arrBeforeFilter = arrBeforeFilter.filter((arr) => {
        return arr.Steps == tempFilterKeys.step;
      });
    }

    if (tempFilterKeys.lesson != "All") {
      arrBeforeFilter = arrBeforeFilter.filter((arr) => {
        return arr.Lesson == tempFilterKeys.lesson;
      });
    }
    columnSortArr = arrBeforeFilter;
    setPadpData([...columnSortArr]);
    setPadpFilters({ ...tempFilterKeys });
  };
  const overallStatus = () => {
    if (padpData.every((data) => data.Status == "Completed")) {
      return <div className={padpStatusStyles.completed}>Completed</div>;
    } else if (padpData.every((data) => data.Status == "Scheduled")) {
      return <div className={padpStatusStyles.scheduled}>Scheduled</div>;
    } else if (padpData.every((data) => data.Status == "On schedule")) {
      return <div className={padpStatusStyles.onSchedule}>On schedule</div>;
    } else if (padpData.every((data) => data.Status == "Behind schedule")) {
      return (
        <div className={padpStatusStyles.behindScheduled}>Behind schedule</div>
      );
    } else {
      return <div className={padpStatusStyles.scheduled}>Scheduled</div>;
    }
  };
  const overallPlannedHours = () => {
    let ph = 0;
    if (padpData.length > 0) {
      padpData.forEach((data) => {
        ph += data.PH ? data.PH : 0;
      });
    }
    return ph;
  };
  const overallActualHours = () => {
    let ah = 0;
    if (padpData.length > 0) {
      padpData.forEach((data) => {
        ah += data.AH ? data.AH : 0;
      });
    }
    return ah;
  };
  const _onColumnClick = (
    ev: React.MouseEvent<HTMLElement>,
    column: IColumn
  ): void => {
    const tempapColumns = padpColumns;
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

    const newDisplayData = _copyAndSort(
      columnSortArr,
      currColumn.fieldName!,
      currColumn.isSortedDescending
    );
    const newMasterData = _copyAndSort(
      columnSortArr,
      currColumn.fieldName!,
      currColumn.isSortedDescending
    );
    setPadpData([...newDisplayData]);
    setPadpMasterData([...newMasterData]);
  };
  function _copyAndSort<T>(
    items: T[],
    columnKey: string,
    isSortedDescending?: boolean
  ): T[] {
    let key = columnKey as keyof T;
    if (key == "Developer") {
      const ascSortFunction = (a, b) => {
        if (a.Developer.name < b.Developer.name) {
          return -1;
        }
        if (a.Developer.name > b.Developer.name) {
          return 1;
        }
        return 0;
      };
      const decSortFunction = (b, a) => {
        if (a.Developer.name < b.Developer.name) {
          return -1;
        }
        if (a.Developer.name > b.Developer.name) {
          return 1;
        }
        return 0;
      };

      return items.sort(isSortedDescending ? ascSortFunction : decSortFunction);
    } else {
      return items
        .slice(0)
        .sort((a: T, b: T) =>
          (isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1
        );
    }
  }
  const ErrorPopup = () => (
    alertify.set("notifier", "position", "top-right"),
    alertify.error("Something when error, please contact system admin.")
  );
  const padpErrorFunction = (error: any, functionName: string) => {
    console.log(error);

    let response = {
      ComponentName: "Product activity delivery plan",
      FunctionName: functionName,
      ErrorMessage: JSON.stringify(error["message"]),
      Title: loggeduseremail,
    };

    Service.SPAddItem({ Listname: "Error Log", RequestJSON: response }).then(
      () => {
        setPadpLoader("noLoader");
        ErrorPopup();
      }
    );
  };

  //Function-Section Ends

  useEffect(() => {
    setPadpLoader("startUpLoader");
    padpGetData();
    getActivityPlanItem();
  }, [padpReRender]);
  return (
    <>
      <div style={{ padding: "5px 15px" }}>
        {padpLoader == "startUpLoader" ? <CustomLoader /> : null}
        {/* Header-Section Starts */}
        <div
          className={styles.padpHeaderSection}
          style={{ paddingBottom: "0" }}
        >
          {/* Title-Section Starts */}
          <div className={styles.padpHeader} style={{ marginBottom: "15px" }}>
            <div className={styles.dpTitle}>
              <Icon
                iconName="NavigateBack"
                className={padpIconStyleClass.navArrow}
                onClick={() => {
                  props.handleclick("ProductActivityPlan", activityPlan_ID);
                }}
              />
              <Label style={{ color: "#000", fontSize: 24, padding: 0 }}>
                Product Activity Delivery plan
              </Label>
            </div>
          </div>
          {/* Title-Section Ends */}
          {/* ATPDetails-Section Starts */}
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
                  {activtyPlanItem.length > 0 ? activtyPlanItem[0].Project : ""}
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
                  {activtyPlanItem.length > 0 ? activtyPlanItem[0].Product : ""}
                </Label>
              </div>
              {padpData.length > 0 ? (
                <div
                  style={{
                    display: "flex",
                    alignItems: "center",
                    justifyContent: "center",
                    marginRight: "15px",
                  }}
                >
                  <Label className={paplabelStyles.titleLabel}>Status :</Label>
                  <Label className={paplabelStyles.labelValue}>
                    {overallStatus()}
                  </Label>
                </div>
              ) : null}
              {/* <div style={{ margin: "0 25px 0 10px" }}>
                <Label style={{ marginRight: 5 }}>Type :</Label>
                <Label style={{ color: "#038387" }}>
                  {activtyPlanItem.length > 0 ? activtyPlanItem[0].Types : ""}
                </Label>
              </div>
              <div style={{ margin: "0 25px 0 10px" }}>
                <Label style={{ marginRight: 5 }}>Project :</Label>
                <Label style={{ color: "#038387" }}>
                  {activtyPlanItem.length > 0 ? activtyPlanItem[0].Project : ""}
                </Label>
              </div>
              <div style={{ margin: "0 25px 0 10px" }}>
                <Label style={{ marginRight: 5 }}>AH/PH :</Label>
                <Label style={{ color: "#038387" }}>
                  {overallActualHours()}/{overallPlannedHours()}
                </Label>
              </div> */}
            </div>
          </div>
          {/* ATPDetails-Section Ends */}
          {/* Filter-Section Starts */}
          <div>
            <div
              style={{
                display: "flex",
                justifyContent: "space-between",
                marginTop: "-5px",
                marginBottom: "10px",
                flexWrap: "wrap",
              }}
            >
              <div
                style={{
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "flex-start",
                  flexWrap: "wrap",
                }}
              >
                <div>
                  <Label styles={padpLabelStyles}>Section</Label>
                  <Dropdown
                    placeholder="Select an option"
                    styles={
                      padpFilters.lesson != "All"
                        ? padpActiveDropdownStyles
                        : padpDropdownStyles
                    }
                    options={padpDropDownOptions.lessonOptns}
                    dropdownWidth={"auto"}
                    onChange={(e, option: any) => {
                      padpListFilter("lesson", option["key"]);
                    }}
                    selectedKey={padpFilters.lesson}
                  />
                </div>
                <div>
                  <Label styles={padpLabelStyles}>Steps</Label>
                  <Dropdown
                    placeholder="Select an option"
                    styles={
                      padpFilters.step != "All"
                        ? padpActiveDropdownStyles
                        : padpDropdownStyles
                    }
                    options={padpDropDownOptions.stepsOptns}
                    dropdownWidth={"auto"}
                    onChange={(e, option: any) => {
                      padpListFilter("step", option["key"]);
                    }}
                    selectedKey={padpFilters.step}
                  />
                </div>
                <div>
                  <Label styles={padpLabelStyles}>Developer</Label>
                  <Dropdown
                    placeholder="Select an option"
                    styles={
                      padpFilters.developer != "All"
                        ? padpActiveDropdownStyles
                        : padpDropdownStyles
                    }
                    options={padpDropDownOptions.developerOptns}
                    dropdownWidth={"auto"}
                    onChange={(e, option: any) => {
                      padpListFilter("developer", option["key"]);
                    }}
                    selectedKey={padpFilters.developer}
                  />
                </div>
                <div>
                  <Label styles={padpLabelStyles}>Status</Label>
                  <Dropdown
                    placeholder="Select an option"
                    styles={
                      padpFilters.status != "All"
                        ? padpActiveDropdownStyles
                        : padpDropdownStyles
                    }
                    options={padpDropDownOptions.statusOptns}
                    dropdownWidth={"auto"}
                    onChange={(e, option: any) => {
                      padpListFilter("status", option["key"]);
                    }}
                    selectedKey={padpFilters.status}
                  />
                </div>
                <div>
                  <Icon
                    iconName="Refresh"
                    title="Click to reset"
                    className={padpIconStyleClass.refresh}
                    onClick={() => {
                      columnSortArr = padpUnsortMasterData;
                      setPadpData(padpUnsortMasterData);
                      columnSortMasterArr = padpUnsortMasterData;
                      setPadpMasterData(padpUnsortMasterData);
                      padpGetAllOptions(padpUnsortMasterData);
                      setPadpFilters({ ...padpFilterKeys });
                      setPadpMasterColumns(padpColumns);
                    }}
                  />
                </div>
              </div>
              <div
                style={{
                  display: "flex",
                  justifyContent: "flex-end",
                  marginLeft: "20px",
                  marginTop: "38px",
                }}
              >
                <div>
                  <Label style={{ marginRight: 5 }}>
                    Number of records :{" "}
                    <span style={{ color: "#038387" }}>{padpData.length}</span>
                  </Label>
                </div>
              </div>
            </div>
          </div>
          {/* Filter-Section Ends */}
        </div>
        {/* Header-Section Ends */}
        {/* Body-Section Starts */}
        <div>
          <div>
            {/* dont remove */}
            <input
              id="forFocus"
              type="text"
              style={{
                width: 0,
                height: 0,
                border: "none",
                position: "absolute",
                top: 0,
                left: 0,
                padding: 0,
              }}
            />
          </div>
          <div
            className={styles.scrollTop}
            onClick={() => {
              document.querySelector("#forFocus")["focus"]();
            }}
          >
            <Icon iconName="Up" style={{ color: "#fff" }} />
          </div>
          <div>
            {/* DetailList-Section Starts */}
            <div>
              {
                <DetailsList
                  items={padpData}
                  columns={padpMasterColumns}
                  setKey="set"
                  layoutMode={DetailsListLayoutMode.justified}
                  selectionMode={SelectionMode.none}
                  // styles={gridStyles}
                  styles={{ root: { width: "100%" } }}
                  onRenderRow={(data, defaultRender) => (
                    <div>
                      {defaultRender({
                        ...data,
                        styles: {
                          root: {
                            background:
                              data.item.IsNew == true
                                ? "linear-gradient(90deg, rgba(250,163,50,0.1491947120645133) 35%, rgba(3,131,135,0.14639359161633403) 100%)"
                                : "#fff",
                            selectors: {
                              "&:hover": {
                                background:
                                  data.item.IsNew == true
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
              }
            </div>
            {/* DetailList-Section Ends */}
            {/* NoData-Section Starts */}
            {padpMasterData.length > 0 ? null : (
              <Label
                style={{
                  paddingLeft: 745,
                  paddingTop: 40,
                }}
                className={padpCommonStyles.inputLabel}
              >
                No Data Found !!!
              </Label>
            )}
            {/* NoData-Section Ends */}
          </div>
        </div>
        {/* Body-Section Ends */}
      </div>
    </>
  );
};

export default ProductActivityDeliveryPlan;
