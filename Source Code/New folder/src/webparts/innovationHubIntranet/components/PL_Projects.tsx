import * as React from "react";
import { useState, useEffect } from "react";
import { Web } from "@pnp/sp/webs";
import {
  DetailsList,
  DetailsListLayoutMode,
  SelectionMode,
  Icon,
  Label,
  Dropdown,
  IDropdownStyles,
  SearchBox,
  ISearchBoxStyles,
  TooltipHost,
  TooltipOverflowMode,
  PrimaryButton,
  IColumn,
} from "@fluentui/react";

import Service from "../components/Services";

import "../ExternalRef/styleSheets/Styles.css";
import styles from "./InnovationHubIntranet.module.scss";
import Pagination from "office-ui-fabric-react-pagination";

import { mergeStyles, mergeStyleSets } from "@fluentui/react/lib/Styling";
import CustomLoader from "./CustomLoader";
import alertify from "alertifyjs";
import "alertifyjs/build/css/alertify.css";
import { arraysEqual } from "office-ui-fabric-react/lib/Utilities";
import * as moment from "moment";

const PL_Projects = (props: any) => {
  let loggeduseremail: string = props.spcontext.pageContext.user.email;
  const sharepointWeb = Web(props.URL);
  const ListNameURL = props.WeblistURL;
  const currentBA = props.PLBAObject.BA;
  const currentSubject = props.PLBAObject.Subject;
  const currentProduct = props.PLBAObject.Product;
  const currentProductVersion = props.PLBAObject.ProductVersion;
  const currentProductId = props.PLBAObject.ProductId;

  let currentpage = 1;
  let totalPageItems = 10;
  let sortPLDataArr = [];
  let sortPLFilterArr = [];

  const PLFilterKey = {
    Project: "",
    TypeOfProject: "All",
    Status: "All",
  };
  const PLDrpDownOptns = {
    TypeOfProject: [{ key: "All", text: "All" }],
    Status: [{ key: "All", text: "All" }],
  };

  const _PLPrjColumns = [
    {
      key: "Column1",
      name: "Deliverable",
      fieldName: "Project",
      minWidth: 250,
      maxWidth: 400,
      onColumnClick: (ev, column) => {
        _onColumnClick(ev, column);
      },
      onRender: (item) => (
        <div>
          <TooltipHost
            content={item.Project}
            overflowMode={TooltipOverflowMode.Parent}
          >
            <span>{item.Project}</span>
          </TooltipHost>
        </div>
      ),
    },

    {
      key: "Column2",
      name: "Version",
      fieldName: "Version",
      minWidth: 100,
      maxWidth: 150,
      onColumnClick: (ev, column) => {
        _onColumnClick(ev, column);
      },
    },
    {
      key: "Column3",
      name: "TOD",
      fieldName: "TypeOfProject",
      minWidth: 100,
      maxWidth: 150,
      onColumnClick: (ev, column) => {
        _onColumnClick(ev, column);
      },
    },
    {
      key: "Column4",
      name: "Status",
      fieldName: "Status",
      minWidth: 150,
      maxWidth: 150,
      onColumnClick: (ev, column) => {
        _onColumnClick(ev, column);
      },
      onRender: (item) => (
        <>
          {item.Status == "Completed" ? (
            <div className={PLStatusStyleClass.completed}>{item.Status}</div>
          ) : item.Status == "Under development" ? (
            <div className={PLStatusStyleClass.scheduled}>{item.Status}</div>
          ) : item.Status == "Not started" ? (
            <div className={PLStatusStyleClass.onSchedule}>{item.Status}</div>
          ) : (
            <div className={PLStatusStyleClass.Onhold}>{item.Status}</div>
          )}
        </>
      ),
    },
    {
      key: "Column5",
      name: "Completion",
      fieldName: "Completion",
      minWidth: 150,
      maxWidth: 250,
      onColumnClick: (ev, column) => {
        _onColumnClick(ev, column);
      },
      onRender: (item) => (
        <div style={{ width: 200, position: "relative" }}>
          {item.Completion >= 75 ? (
            <div className={PLCompletionStyleClass.default}>
              <span className={PLCompletionStyleClass.percentageText}>
                {Math.round(item.Completion)}%
              </span>{" "}
              <div
                style={{ width: `${item.Completion}%` }}
                className={PLCompletionStyleClass.completed}
              >
                {/* {item.Completion} */}
              </div>
            </div>
          ) : item.Completion >= 50 ? (
            <div className={PLCompletionStyleClass.default}>
              <span className={PLCompletionStyleClass.percentageText}>
                {Math.round(item.Completion)}%
              </span>{" "}
              <div
                style={{ width: `${item.Completion}%` }}
                className={PLCompletionStyleClass.scheduled}
              >
                {/* {item.Completion} */}
              </div>
            </div>
          ) : item.Completion >= 25 ? (
            <div className={PLCompletionStyleClass.default}>
              <span className={PLCompletionStyleClass.percentageText}>
                {Math.round(item.Completion)}%
              </span>{" "}
              <div
                style={{ width: `${item.Completion}%` }}
                className={PLCompletionStyleClass.onSchedule}
              >
                {/* {item.Completion} */}
              </div>
            </div>
          ) : item.Completion >= 0 ? (
            <div className={PLCompletionStyleClass.default}>
              <span className={PLCompletionStyleClass.percentageText}>
                {Math.round(item.Completion)}%
              </span>{" "}
              <div
                style={{ width: `${item.Completion}%` }}
                className={PLCompletionStyleClass.behindScheduled}
              >
                {/* {item.Completion} */}
              </div>
            </div>
          ) : (
            ""
          )}
        </div>
      ),
    },
    {
      key: "Column6",
      name: "Link",
      fieldName: "link",
      minWidth: 30,
      maxWidth: 30,

      onRender: (item) => (
        <>
          <Icon
            iconName="Link12"
            className={PLIconStyleClass.link}
            onClick={() => {
              props.selectPLFunction(
                "Resources",
                "Project",
                item.Project,
                item.Version,
                item.ID
              );
            }}
          />
        </>
      ),
    },
  ];
  const PLSearchBoxStyles: Partial<ISearchBoxStyles> = {
    root: {
      width: 200,
      marginRight: "15px",
      backgroundColor: "#F5F5F7",
      outline: "none",
      border: "1px solid #E8E8EA",
      borderRadius: "4px",
      marginTop: "3px",
    },
    icon: { fontSize: 14, color: "#000" },
  };
  const PLActiveSearchBoxStyles: Partial<ISearchBoxStyles> = {
    root: {
      width: 200,
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
  const PLlabelStyles = mergeStyleSets({
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
    navLabels: [
      {
        color: "#2392B2",
        fontSize: "16px",
        cursor: "pointer",
      },
    ],
    navViewLabels: [
      {
        fontSize: "16px",
        cursor: "pointer",
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
        marginTop: 28,
        marginLeft: "10px",
        fontWeight: "500",
        color: "#323130",
        fontSize: "13px",
      },
    ],
  });
  const PLProjectInfo = mergeStyles({
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    margin: "0 15px 0 0",
  });
  const PLIconStyle = mergeStyles({
    fontSize: 17,
    height: 14,
    width: 17,
    cursor: "pointer",
  });
  const PLIconStyleClass = mergeStyleSets({
    link: [
      {
        fontSize: 18,
        height: 14,
        width: 17,
        color: "#2392B2",
        cursor: "pointer",
        marginLeft: "4px",
      },
    ],
    rightArrow: [
      { color: "#2392B2", marginRight: 10, fontSize: 20 },
      PLIconStyle,
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
        marginTop: 31,
        borderRadius: 2,
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        ":hover": {
          backgroundColor: "#025d60",
        },
      },
    ],
    ChevronLeftMed: [
      {
        cursor: "pointer",
        color: "#2392b2",
        fontSize: 24,
        marginTop: 3,
        marginRight: 12,
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
  const PLStatusStyle = mergeStyles({
    textAlign: "center",
    borderRadius: "25px",
    width: 120,
  });
  const PLStatusStyleClass = mergeStyleSets({
    completed: [
      {
        fontWeight: "600",
        padding: "3px",
        color: "#438700",
        backgroundColor: "#D9FFB3",
      },
      PLStatusStyle,
    ],
    scheduled: [
      {
        fontWeight: "600",
        padding: "3px",
        color: "#9f6700",
        backgroundColor: "#FFDB99",
      },
      PLStatusStyle,
    ],
    onSchedule: [
      {
        fontWeight: "600",
        padding: "3px",
        color: "rgb(0 191 198)",
        backgroundColor: "rgb(210 241 241)",
      },
      PLStatusStyle,
    ],
    behindScheduled: [
      {
        fontWeight: "600",
        padding: "3px",
        color: "#FF0000",
        backgroundColor: "#FFB3B3",
      },
      PLStatusStyle,
    ],
    Onhold: [
      {
        fontWeight: "600",
        padding: "3px",
        color: "#773030",
        backgroundColor: "#e6b1b1",
      },
      PLStatusStyle,
    ],
  });
  const PLCompletionStyle = mergeStyles({
    textAlign: "center",
    borderRadius: "25px",
    height: 17,
    width: 200,
  });
  const PLCompletionStyleClass = mergeStyleSets({
    completed: [
      {
        fontWeight: "600",
        color: "#438700",
        backgroundColor: "#D9FFB3",
      },
      PLCompletionStyle,
    ],
    scheduled: [
      {
        fontWeight: "600",
        color: "#9f6700",
        backgroundColor: "#FFDB99",
      },
      PLCompletionStyle,
    ],
    onSchedule: [
      {
        fontWeight: "600",
        color: "#9C9C00; ",
        backgroundColor: "#EEEEAE",
      },
      PLCompletionStyle,
    ],
    behindScheduled: [
      {
        fontWeight: "600",
        color: "#FF0000",
        backgroundColor: "#FFB3B3",
      },
      PLCompletionStyle,
    ],
    default: [
      {
        fontWeight: "600",
        position: "relative",
        backgroundColor: "#edebe9",
      },
      PLCompletionStyle,
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
  const PLDropdownStyles: Partial<IDropdownStyles> = {
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
    callout: {
      maxHeight: "300px",
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
  const PLActiveDropdownStyles: Partial<IDropdownStyles> = {
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
    callout: {
      maxHeight: "300px",
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

  // Use State
  const [PLPrjReRender, setPLPrjReRender] = useState(false);
  const [PLPrjMaster, setPLPrjMaster] = useState([]);
  const [PLPrjData, setPLPrjData] = useState([]);
  const [PLPrjFilData, setPLPrjFilData] = useState([]);
  const [PLPrjDisplay, setPLPrjDisplay] = useState([]);
  const [PLPrjFilterKey, setPLPrjFilterKey] = useState(PLFilterKey);
  const [PLPrjDrpDownOptns, setPLPrjDrpDownOptns] = useState(PLDrpDownOptns);
  const [PLPrjLoader, setPLPrjLoader] = useState("noLoader");
  const [PLPrjColumns, setPLPrjColumns] = useState(_PLPrjColumns);
  const [PLPrjcurPage, setPLPrjCurPage] = useState(currentpage);

  const getProjects = () => {
    let _apCurrentData = [];
    sharepointWeb.lists
      .getByTitle(ListNameURL)
      .items.filter(
        `BusinessArea eq '${currentBA}' and Master_x0020_Project/Id eq '${currentProductId}'`
      )
      .select(
        "*",
        "Master_x0020_Project/Title",
        "Master_x0020_Project/Id",
        "Master_x0020_Project/ProductVersion"
      )
      .expand("Master_x0020_Project")
      .top(5000)
      .orderBy("Modified", false)
      .get()
      .then((items) => {
        items.forEach((item) => {
          _apCurrentData.push({
            ID: item.ID,
            ProjectType: item.ProjectType,
            Title: item.Title,
            ProjectVersion: item.ProjectVersion ? item.ProjectVersion : "V1",
            ProductId: item.Master_x0020_ProjectId,
            ProductName: item.Master_x0020_Project
              ? item.Master_x0020_Project.Title
              : "",
            ProductVersion: item.Master_x0020_Project
              ? item.Master_x0020_Project.ProductVersion
                ? item.Master_x0020_Project.ProductVersion
                : "V1"
              : "V1",
          });
        });

        getATPDetails(_apCurrentData);
        // getDeliveryPlan(items);
      })
      .catch((err) => {
        ErrorFunction(err, "getApData");
      });
  };
  const getATPDetails = (records) => {
    let _ATPArr = [];
    sharepointWeb.lists
      .getByTitle("Activity Plan")
      .items.filter(`Product eq '${currentProduct}'`)
      .select("*", "FieldValuesAsText/CompletedDate")
      .expand("FieldValuesAsText")
      .orderBy("Modified", false)
      .top(5000)
      .get()
      .then((items) => {
        items.forEach((item) => {
          _ATPArr.push({
            ID: item.ID,
            ProjectType: item.ProjectType,
            Project: item.Project,
            ProjectVersion: item.ProjectVersion ? item.ProjectVersion : "V1",
            Product: item.Product ? item.Product : "",
            ProductVersion: item.ProductVersion ? item.ProductVersion : "V1",
            Status: item.Status,
            Completion: item.Completion,
          });
        });
        dataManipulation(records, _ATPArr);
      })
      .catch((err) => {
        console.log(err);
      });
  };

  const dataManipulation = (APData, ATPData) => {
    let _projectArr = [];
    let count = 0;

    if (APData.length > 0) {
      APData.forEach((ap) => {
        count++;
        let curATPDetails = ATPData.filter((arr) => {
          return (
            arr.Project == ap.Title &&
            arr.ProjectVersion == ap.ProjectVersion &&
            arr.ProductVersion == currentProductVersion
          );
        });

        let CompletedDetails = curATPDetails.filter((arr) => {
          return arr.Status == "Completed";
        });

        let varCompletion = 0;
        curATPDetails.forEach((item) => {
          varCompletion += item.Completion;
        });
        varCompletion =
          varCompletion > 0 && curATPDetails.length > 0
            ? varCompletion / curATPDetails.length
            : 0;

        _projectArr.push({
          ID: ap.ID,
          Project: ap.Title ? ap.Title : "",
          Version: ap.ProjectVersion ? ap.ProjectVersion : "V1",
          TypeOfProject: ap.ProjectType ? ap.ProjectType : "",
          Status:
            varCompletion > 0
              ? CompletedDetails.length == curATPDetails.length &&
                CompletedDetails.length != 0 &&
                curATPDetails.length != 0
                ? "Completed"
                : "Under development"
              : "Not started",
          // CompletedDetails.length != 0 && ATPData.length != 0
          //   ? CompletedDetails.length == ATPData.length
          //     ? "Completed"
          //     : "Under development"
          //   : "Not started",
          Completion: varCompletion,
        });

        if (
          PLDrpDownOptns.TypeOfProject.findIndex((_productDrpDwn) => {
            return _productDrpDwn.key == ap.ProjectType;
          }) == -1 &&
          ap.ProjectType
        ) {
          PLDrpDownOptns.TypeOfProject.push({
            key: ap.ProjectType,
            text: ap.ProjectType,
          });
        }

        // if (
        //   PLDrpDownOptns.Status.findIndex((_projectDrpDwn) => {
        //     return _projectDrpDwn.key == ap.Status;
        //   }) == -1 &&
        //   ap.Status
        // ) {
        //   PLDrpDownOptns.Status.push({
        //     key: ap.Status,
        //     text: ap.Status,
        //   });
        // }

        if (APData.length == count) {
          PLDrpDownOptns.Status.push(
            {
              key: "Not started",
              text: "Not started",
            },
            {
              key: "Under development",
              text: "Under development",
            },
            {
              key: "Completed",
              text: "Completed",
            }
          );

          setPLPrjDrpDownOptns({ ...PLDrpDownOptns });
          setPLPrjMaster([..._projectArr]);
          setPLPrjData([..._projectArr]);
          sortPLDataArr = [..._projectArr];
          setPLPrjFilData([..._projectArr]);
          sortPLFilterArr = [..._projectArr];
          paginate(1, [..._projectArr]);
          setPLPrjLoader("noLoader");
        }
      });
    } else {
      setPLPrjDrpDownOptns({ ...PLDrpDownOptns });
      setPLPrjMaster([..._projectArr]);
      setPLPrjData([..._projectArr]);
      sortPLDataArr = [..._projectArr];
      setPLPrjFilData([..._projectArr]);
      sortPLFilterArr = [..._projectArr];
      paginate(1, [..._projectArr]);
      setPLPrjLoader("noLoader");
    }
  };
  const getDeliveryPlan = (records) => {
    let _projectArr = [];
    let count = 0;
    records.forEach((record) => {
      sharepointWeb.lists
        .getByTitle("Delivery Plan")
        .items.filter("AnnualPlanID eq '" + record.ID + "' ")
        .top(5000)
        .get()
        .then((items) => {
          count++;
          let completedLength = items.filter((arr) => {
            return arr.Status == "Completed";
          }).length;

          _projectArr.push({
            ID: record.ID,
            Project: record.Title ? record.Title : "",
            Version: record.ProjectVersion ? record.ProjectVersion : "V1",
            TypeOfProject: record.ProjectType ? record.ProjectType : "",
            Status: record.Status,
            Completion:
              record.Status == "Completed"
                ? 100
                : completedLength != 0 && items.length != 0
                ? (completedLength / items.length) * 100
                : 0,
          });

          if (
            PLDrpDownOptns.TypeOfProject.findIndex((_productDrpDwn) => {
              return _productDrpDwn.key == record.ProjectType;
            }) == -1 &&
            record.ProjectType
          ) {
            PLDrpDownOptns.TypeOfProject.push({
              key: record.ProjectType,
              text: record.ProjectType,
            });
          }

          if (
            PLDrpDownOptns.Status.findIndex((_projectDrpDwn) => {
              return _projectDrpDwn.key == record.Status;
            }) == -1 &&
            record.Status
          ) {
            PLDrpDownOptns.Status.push({
              key: record.Status,
              text: record.Status,
            });
          }

          if (records.length == count) {
            setPLPrjDrpDownOptns({ ...PLDrpDownOptns });
            setPLPrjMaster([..._projectArr]);
            setPLPrjData([..._projectArr]);
            sortPLDataArr = [..._projectArr];
            setPLPrjFilData([..._projectArr]);
            sortPLFilterArr = [..._projectArr];
            paginate(1, [..._projectArr]);
            setPLPrjLoader("noLoader");
          }
        });
    });
  };
  const paginate = (pagenumber, data) => {
    if (data.length > 0) {
      let lastIndex: number = pagenumber * totalPageItems;
      let firstIndex: number = lastIndex - totalPageItems;
      let paginatedItems = data.slice(firstIndex, lastIndex);
      currentpage = pagenumber;
      setPLPrjDisplay(paginatedItems);
      setPLPrjCurPage(pagenumber);
    } else {
      setPLPrjDisplay([]);
      setPLPrjCurPage(1);
    }
  };
  const PLFilterFunction = (key, value) => {
    let tempFilterKey = { ...PLPrjFilterKey };
    tempFilterKey[key] = value;

    let tempArr = [...PLPrjData];

    if (tempFilterKey.Project != "") {
      tempArr = tempArr.filter((arr) => {
        return arr.Project.toLowerCase().includes(
          tempFilterKey.Project.toLowerCase()
        );
      });
    }
    if (tempFilterKey.TypeOfProject != "All") {
      tempArr = tempArr.filter((arr) => {
        return arr.TypeOfProject == tempFilterKey.TypeOfProject;
      });
    }
    if (tempFilterKey.Status != "All") {
      tempArr = tempArr.filter((arr) => {
        return arr.Status == tempFilterKey.Status;
      });
    }
    setPLPrjFilData([...tempArr]);
    sortPLFilterArr = [...tempArr];
    paginate(1, [...tempArr]);
    setPLPrjFilterKey({ ...tempFilterKey });
  };
  const ErrorFunction = (error: any, functionName: string) => {
    console.log(error);
    setPLPrjLoader("noLoader");

    // let response = {
    //   ComponentName: "PL_Projects",
    //   FunctionName: functionName,
    //   ErrorMessage: JSON.stringify(error["message"]),
    //   Title: loggeduseremail,
    // };

    // Service.SPAddItem({ Listname: "Error Log", RequestJSON: response }).then(
    //   () => {
    ErrorPopup();
    //   }
    // );
  };
  const ErrorPopup = () => (
    alertify.set("notifier", "position", "top-right"),
    alertify.error("Something when error, please contact system admin.")
  );

  // Sorting Function
  const _onColumnClick = (
    ev: React.MouseEvent<HTMLElement>,
    column: IColumn
  ): void => {
    const tempapColumns = _PLPrjColumns;
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

    const newPLDataArr = _copyAndSort(
      sortPLDataArr,
      currColumn.fieldName!,
      currColumn.isSortedDescending
    );
    const newPLFilterArr = _copyAndSort(
      sortPLFilterArr,
      currColumn.fieldName!,
      currColumn.isSortedDescending
    );
    setPLPrjData([...newPLDataArr]);
    setPLPrjFilData([...newPLFilterArr]);
    paginate(1, newPLFilterArr);
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

  //Use Effect
  useEffect(() => {
    setPLPrjLoader("startUpLoader");
    getProjects();
  }, [PLPrjReRender]);

  return (
    <>
      <div style={{ padding: "5px 15px" }}>
        {PLPrjLoader == "startUpLoader" ? (
          <CustomLoader />
        ) : (
          <>
            <div
              style={{
                display: "flex",
                marginBottom: 10,
              }}
            >
              <div
                style={{
                  display: "flex",
                  alignItems: "center",
                }}
                onClick={() => {
                  props.selectPLFunction("BusinessArea");
                }}
              >
                <Label className={PLlabelStyles.navLabels}>Home</Label>
                <Icon
                  iconName="ChevronRight"
                  title="Click to navigate"
                  className={PLIconStyleClass.rightArrow}
                />
              </div>
              <div
                style={{
                  display: "flex",
                  alignItems: "center",
                }}
                onClick={() => {
                  props.selectPLFunction("Subject");
                }}
              >
                <Label className={PLlabelStyles.navLabels}>{currentBA}</Label>
                <Icon
                  iconName="ChevronRight"
                  title="Click to navigate"
                  className={PLIconStyleClass.rightArrow}
                />
              </div>
              <div
                style={{
                  display: "flex",
                  alignItems: "center",
                }}
                onClick={() => {
                  props.selectPLFunction("Product");
                }}
              >
                <Label className={PLlabelStyles.navLabels}>
                  {currentSubject}
                </Label>
                <Icon
                  iconName="ChevronRight"
                  title="Click to navigate"
                  className={PLIconStyleClass.rightArrow}
                />
              </div>
              <div>
                <Label className={PLlabelStyles.navViewLabels}>
                  {currentProduct + " " + currentProductVersion}
                </Label>
              </div>
            </div>
            <div
              style={{
                display: "flex",
                alignItems: "flex-start",
                justifyContent: "space-between",
                marginBottom: 10,
              }}
            >
              <div
                style={{
                  display: "flex",
                  alignItems: "center",
                }}
              >
                <Icon
                  aria-label="ChevronLeftMed"
                  iconName="NavigateBack"
                  className={PLIconStyleClass.ChevronLeftMed}
                  onClick={() => {
                    props.selectPLFunction("Product");
                  }}
                />

                <div className={styles.dpTitle}>
                  <Label style={{ fontSize: 24, padding: 0 }}>
                    Deliverable
                  </Label>
                </div>
              </div>
            </div>
            <div
              style={{
                display: "flex",
                marginTop: 15,
                justifyContent: "space-between",
              }}
            >
              <div className={styles.Section1}>
                <div className={PLProjectInfo}>
                  <Label className={PLlabelStyles.titleLabel}>
                    Business area :
                  </Label>
                  <Label className={PLlabelStyles.labelValue}>
                    {currentBA}
                  </Label>
                </div>
                <div className={PLProjectInfo}>
                  <Label className={PLlabelStyles.titleLabel}>Subject :</Label>
                  <Label className={PLlabelStyles.labelValue}>
                    {currentSubject}
                  </Label>
                </div>
                <div className={PLProjectInfo}>
                  <Label className={PLlabelStyles.titleLabel}>Product :</Label>
                  <Label className={PLlabelStyles.labelValue}>
                    {currentProduct + " " + currentProductVersion}
                  </Label>
                </div>
              </div>
            </div>
            <div
              style={{
                display: "flex",
                alignItems: "center",
                justifyContent: "space-between",
              }}
            >
              <div
                style={{
                  display: "flex",
                  alignItems: "center",
                  flexWrap: "wrap",
                }}
              >
                <div>
                  <Label className={PLlabelStyles.inputLabels}>
                    Deliverable
                  </Label>
                  <SearchBox
                    styles={
                      PLPrjFilterKey.Project
                        ? PLActiveSearchBoxStyles
                        : PLSearchBoxStyles
                    }
                    value={PLPrjFilterKey.Project}
                    onChange={(e, value) => {
                      PLFilterFunction("Project", value);
                    }}
                  />
                </div>
                <div>
                  <Label className={PLlabelStyles.inputLabels}>TOD</Label>
                  <Dropdown
                    placeholder="Select an option"
                    styles={
                      PLPrjFilterKey.TypeOfProject != "All"
                        ? PLActiveDropdownStyles
                        : PLDropdownStyles
                    }
                    options={PLPrjDrpDownOptns.TypeOfProject}
                    dropdownWidth={"auto"}
                    selectedKey={PLPrjFilterKey.TypeOfProject}
                    onChange={(e, option: any) => {
                      PLFilterFunction("TypeOfProject", option["key"]);
                    }}
                  />
                </div>
                <div>
                  <Label className={PLlabelStyles.inputLabels}>Status</Label>
                  <Dropdown
                    placeholder="Select an option"
                    styles={
                      PLPrjFilterKey.Status != "All"
                        ? PLActiveDropdownStyles
                        : PLDropdownStyles
                    }
                    options={PLPrjDrpDownOptns.Status}
                    dropdownWidth={"auto"}
                    selectedKey={PLPrjFilterKey.Status}
                    onChange={(e, option: any) => {
                      PLFilterFunction("Status", option["key"]);
                    }}
                  />
                </div>
                <div>
                  <Icon
                    iconName="Refresh"
                    title="Click to reset"
                    className={PLIconStyleClass.refresh}
                    onClick={() => {
                      setPLPrjData([...PLPrjMaster]);
                      sortPLDataArr = [...PLPrjMaster];
                      setPLPrjFilData([...PLPrjMaster]);
                      sortPLFilterArr = [...PLPrjMaster];
                      paginate(1, [...PLPrjMaster]);
                      setPLPrjFilterKey(PLFilterKey);
                      setPLPrjColumns(_PLPrjColumns);
                    }}
                  />
                </div>
              </div>
              <div>
                <Label className={PLlabelStyles.NORLabel}>
                  Number of records:{" "}
                  <b style={{ color: "#038387" }}>{PLPrjFilData.length}</b>
                </Label>
              </div>
            </div>
            <div style={{ marginTop: "10px" }}>
              <DetailsList
                items={PLPrjDisplay}
                columns={PLPrjColumns}
                setKey="set"
                layoutMode={DetailsListLayoutMode.justified}
                selectionMode={SelectionMode.none}
                // styles={gridStyles}
                styles={{ root: { width: "100%" } }}
              />
              <div
                style={{
                  display: "flex",
                  justifyContent: "center",
                  margin: "20px 0",
                }}
              >
                {PLPrjFilData.length > 0 ? (
                  <Pagination
                    currentPage={PLPrjcurPage}
                    totalPages={
                      PLPrjFilData.length > 0
                        ? Math.ceil(PLPrjFilData.length / totalPageItems)
                        : 1
                    }
                    onChange={(page) => {
                      paginate(page, PLPrjFilData);
                    }}
                  />
                ) : (
                  <div
                    style={{
                      display: "flex",
                      justifyContent: "center",
                      marginTop: "15px",
                    }}
                  >
                    <Label style={{ color: "#2392B2" }}>
                      No data Found !!!
                    </Label>
                  </div>
                )}
              </div>
            </div>
          </>
        )}
      </div>
    </>
  );
};
export default PL_Projects;
