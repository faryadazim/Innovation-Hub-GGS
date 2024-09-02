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

const PL_Subject = (props: any) => {
  let loggeduseremail: string = props.spcontext.pageContext.user.email;
  const sharepointWeb = Web(props.URL);
  const ListNameURL = props.WeblistURL;
  const currentBA = props.PLBAObject.BA;
  let currentpage = 1;
  let totalPageItems = 10;
  let sortPLDataArr = [];
  let sortPLFilterArr = [];

  const PLFilterKey = {
    Subject: "",
  };

  const _PLSubColumns = [
    {
      key: "Column1",
      name: "Subject",
      fieldName: "Subject",
      minWidth: 300,
      maxWidth: 400,
      onColumnClick: (ev, column) => {
        _onColumnClick(ev, column);
      },
      onRender: (item) => (
        <div>
          <TooltipHost
            content={item.Subject}
            overflowMode={TooltipOverflowMode.Parent}
          >
            <span>{item.Subject}</span>
          </TooltipHost>
        </div>
      ),
    },
    {
      key: "Column2",
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
                "Product",
                "Subject",
                item.Subject,
                "",
                ""
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

  // Use State
  const [PLSubReRender, setPLSubReRender] = useState(false);
  const [PLSubMaster, setPLSubMaster] = useState([]);
  const [PLSubData, setPLSubData] = useState([]);
  const [PLSubFilData, setPLSubFilData] = useState([]);
  const [PLSubDisplay, setPLSubDisplay] = useState([]);
  const [PLSubFilterKey, setPLSubFilterKey] = useState(PLFilterKey);
  const [PLSubLoader, setPLSubLoader] = useState("noLoader");
  const [PLSubColumns, setPLSubColumns] = useState(_PLSubColumns);
  const [PLSubcurPage, setPLSubCurPage] = useState(currentpage);

  const getSubjects = () => {
    let _subjectArr = [];
    sharepointWeb.lists
      .getByTitle(ListNameURL)
      .items.filter(`BusinessArea eq '${currentBA}'`)
      .select("*", "Master_x0020_Project/Subject")
      .expand("Master_x0020_Project")
      .top(5000)
      .orderBy("Modified", false)
      .get()
      .then((items) => {
        items.forEach((item: any) => {
          if (
            item.Master_x0020_Project &&
            _subjectArr.findIndex((arr) => {
              return arr.Subject == item.Master_x0020_Project.Subject;
            }) == -1 &&
            item.Master_x0020_Project.Subject
          ) {
            _subjectArr.push({
              Subject: item.Master_x0020_ProjectId
                ? item.Master_x0020_Project.Subject
                : "",
            });
          }
        });
        setPLSubMaster(_subjectArr);
        setPLSubData([..._subjectArr]);
        sortPLDataArr = _subjectArr;
        setPLSubFilData([..._subjectArr]);
        sortPLFilterArr = [..._subjectArr];
        paginate(1, _subjectArr);
        setPLSubLoader("noLoader");
      })
      .catch((err) => {
        ErrorFunction(err, "getApData");
      });
  };

  const paginate = (pagenumber, data) => {
    if (data.length > 0) {
      let lastIndex: number = pagenumber * totalPageItems;
      let firstIndex: number = lastIndex - totalPageItems;
      let paginatedItems = data.slice(firstIndex, lastIndex);
      currentpage = pagenumber;
      setPLSubDisplay(paginatedItems);
      setPLSubCurPage(pagenumber);
    } else {
      setPLSubDisplay([]);
      setPLSubCurPage(1);
    }
  };

  const PLFilterFunction = (key, value) => {
    let tempFilterKey = { ...PLSubFilterKey };
    tempFilterKey[key] = value;

    let tempArr = [...PLSubData];

    if (tempFilterKey.Subject != "") {
      tempArr = tempArr.filter((arr) => {
        return arr.Subject.toLowerCase().includes(
          tempFilterKey.Subject.toLowerCase()
        );
      });
    }
    setPLSubFilData([...tempArr]);
    sortPLFilterArr = [...tempArr];
    paginate(1, [...tempArr]);
    setPLSubFilterKey({ ...tempFilterKey });
  };

  // Sorting Function
  const _onColumnClick = (
    ev: React.MouseEvent<HTMLElement>,
    column: IColumn
  ): void => {
    const tempapColumns = _PLSubColumns;
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
    setPLSubData([...newPLDataArr]);
    setPLSubFilData([...newPLFilterArr]);
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

  const ErrorFunction = (error: any, functionName: string) => {
    console.log(error);
    setPLSubLoader("noLoader");

    // let response = {
    //   ComponentName: "PL_Subject",
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

  //Use Effect
  useEffect(() => {
    setPLSubLoader("startUpLoader");
    getSubjects();
  }, [PLSubReRender]);

  return (
    <>
      <div style={{ padding: "5px 15px" }}>
        {PLSubLoader == "startUpLoader" ? (
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
              <div>
                <Label className={PLlabelStyles.navViewLabels}>
                  {currentBA}
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
                    props.selectPLFunction("BusinessArea");
                  }}
                />

                <div className={styles.dpTitle}>
                  <Label style={{ fontSize: 24, padding: 0 }}>Subjects</Label>
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
                  <Label className={PLlabelStyles.inputLabels}>Subject</Label>
                  <SearchBox
                    styles={
                      PLSubFilterKey.Subject
                        ? PLActiveSearchBoxStyles
                        : PLSearchBoxStyles
                    }
                    value={PLSubFilterKey.Subject}
                    onChange={(e, value) => {
                      PLFilterFunction("Subject", value);
                    }}
                  />
                </div>
                <div>
                  <Icon
                    iconName="Refresh"
                    title="Click to reset"
                    className={PLIconStyleClass.refresh}
                    onClick={() => {
                      paginate(1, [...PLSubMaster]);
                      setPLSubFilterKey(PLFilterKey);
                      setPLSubColumns(_PLSubColumns);
                    }}
                  />
                </div>
              </div>
              <div>
                <Label className={PLlabelStyles.NORLabel}>
                  Number of records:{" "}
                  <b style={{ color: "#038387" }}>{PLSubFilData.length}</b>
                </Label>
              </div>
            </div>
            <div style={{ marginTop: "10px" }}>
              <DetailsList
                items={PLSubDisplay}
                columns={PLSubColumns}
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
                {PLSubFilData.length > 0 ? (
                  <Pagination
                    currentPage={PLSubcurPage}
                    totalPages={
                      PLSubFilData.length > 0
                        ? Math.ceil(PLSubFilData.length / totalPageItems)
                        : 1
                    }
                    onChange={(page) => {
                      paginate(page, PLSubFilData);
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
export default PL_Subject;
