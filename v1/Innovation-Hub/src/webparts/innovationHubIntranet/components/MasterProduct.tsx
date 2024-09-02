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
  SearchBox,
  ISearchBoxStyles,
  Modal,
  IModalStyles,
  TextField,
  ITextFieldStyles,
  Spinner,
  Persona,
  PersonaPresence,
  PersonaSize,
  IColumn,
} from "@fluentui/react";

import Service from "../components/Services";

import "../ExternalRef/styleSheets/Styles.css";
import Pagination from "office-ui-fabric-react-pagination";
import { mergeStyles, mergeStyleSets } from "@fluentui/react/lib/Styling";
import styles from "./InnovationHubIntranet.module.scss";
import CustomLoader from "./CustomLoader";
import alertify from "alertifyjs";
import "alertifyjs/build/css/alertify.css";
import { WebPartContext } from "@microsoft/sp-webpart-base";

interface IProps {
  context: WebPartContext;
  spcontext: any;
  graphContent: any;
  URL: string;
  handleclick: any;
  pageType: string;
  peopleList: any;
}
interface IData {
  ID: number;
  Program: string;
  ProductVersion: string;
  Subject: string;
  Created: Date;
  Modified: Date;
  CreatedBy: string;
  ModifiedBy: string;
  CreatedByName: string;
  ModifiedByName: string;
}
interface IFilters {
  program: string;
}

let columnSortArr: IData[] = [];
let columnSortMasterArr: IData[] = [];

const MasterProduct = (props: IProps): JSX.Element => {
  // Variable-Declaration-Section Starts
  const sharepointWeb: any = Web(props.URL);
  const ListName: string = "Master Product List";
  let loggeduseremail: string = props.spcontext.pageContext.user.email;
  let loggedusername: string = props.spcontext.pageContext.user.displayName;

  const mpAllItems: IData[] = [];
  const mpColumns: IColumn[] = [
    // {
    //   key: "ID",
    //   name: "ID",
    //   fieldName: "ID",
    //   minWidth: 50,
    //   maxWidth: 50,
    // },
    {
      key: "Program",
      name: "Product",
      fieldName: "Program",
      minWidth: 100,
      maxWidth: 300,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
      onRender: (item): string =>
        item.Program + " " + (item.ProductVersion ? item.ProductVersion : "V1"),
    },
    {
      key: "Subject",
      name: "Subject",
      fieldName: "Subject",
      minWidth: 100,
      maxWidth: 300,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
    },
    {
      key: "Created",
      name: "Created",
      fieldName: "Created",
      minWidth: 50,
      maxWidth: 100,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
      onRender: (item: IData): string =>
        moment(item.Created).format("DD/MM/YYYY"),
    },
    // {
    //   key: "Modified",
    //   name: "Modified",
    //   fieldName: "Modified",
    //   minWidth: 50,
    //   maxWidth: 100,
    //   onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
    //     _onColumnClick(ev, column);
    //   },
    //   onRender: (item: IData): string =>
    //     moment(item.Modified).format("DD/MM/YYYY"),
    // },
    {
      key: "CreatedBy",
      name: "Created By",
      fieldName: "CreatedBy",
      minWidth: 125,
      maxWidth: 300,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
      onRender: (item: IData): JSX.Element => (
        <div style={{ display: "flex" }}>
          <div
            style={{
              marginTop: "-6px",
            }}
            title={item.CreatedByName}
          >
            <Persona
              size={PersonaSize.size32}
              presence={PersonaPresence.none}
              imageUrl={
                "/_layouts/15/userphoto.aspx?size=S&username=" +
                `${item.CreatedBy}`
              }
            />
          </div>
          <div>
            <span style={{ fontSize: "13px" }}>{item.CreatedByName}</span>
          </div>
        </div>
      ),
    },
    // {
    //   key: "ModifiedBy",
    //   name: "Modified By",
    //   fieldName: "ModifiedBy",
    //   minWidth: 125,
    //   maxWidth: 300,
    //   onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
    //     _onColumnClick(ev, column);
    //   },
    //   onRender: (item: IData): JSX.Element => (
    //     <div style={{ display: "flex" }}>
    //       <div
    //         style={{
    //           marginTop: "-6px",
    //         }}
    //         title={item.ModifiedByName}
    //       >
    //         <Persona
    //           size={PersonaSize.size32}
    //           presence={PersonaPresence.none}
    //           imageUrl={
    //             "/_layouts/15/userphoto.aspx?size=S&username=" +
    //             `${item.ModifiedBy}`
    //           }
    //         />
    //       </div>
    //       <div>
    //         <span style={{ fontSize: "13px" }}>{item.ModifiedByName}</span>
    //       </div>
    //     </div>
    //   ),
    // },
    {
      key: "Action",
      name: "Action",
      fieldName: "Action",
      minWidth: 70,
      maxWidth: 150,

      onRender: (item: IData): JSX.Element => (
        <>
          <Icon
            iconName="Edit"
            title="Edit product"
            className={mpIconStyleClass.edit}
            onClick={() => {
              setProdSubject(item.Subject);
              setPopup({ type: "edit", selectedItem: [{ ...item }] });
              setProductValue(item.Program);
              setproductVersionValue(item.ProductVersion);
            }}
          />
          <Icon
            iconName="Delete"
            className={mpIconStyleClass.delete}
            onClick={() => {
              setPopup({ type: "delete", selectedItem: [{ ...item }] });
            }}
          />
        </>
      ),
    },
  ];
  const mpFilterKeys: IFilters = { program: "" };
  const mpBigiconStyleClass = mergeStyleSets({
    ChevronLeftMed: [
      {
        cursor: "pointer",
        color: "#2392b2",
        fontSize: 24,
        marginTop: "3px",
        marginRight: 15,
      },
    ],
  });
  let currentpage: number = 1;
  let totalPageItems: number = 10;
  // Variable-Declaration-Section Ends
  // Style-Section Starts
  const mpDetailsListStyles: Partial<IDetailsListStyles> = {
    root: {},
    headerWrapper: {},
    contentWrapper: {
      ".ms-DetailsRow-cell": {
        paddingBottom: "0 !important",
      },
    },
  };
  const mpLabelStyles: Partial<ILabelStyles> = {
    root: {
      width: 165,
      marginTop: 5,
      marginRight: 10,
      fontSize: 13,
      color: "#323130",
    },
  };
  const mpSearchBoxStyles: Partial<ISearchBoxStyles> = {
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
  const mpActiveSearchBoxStyles: Partial<ISearchBoxStyles> = {
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
  const mpModalStyles: Partial<IModalStyles> = {
    root: { borderRadius: "none" },
    main: {
      width: 500,
      minHeight: 250,
      margin: 10,
      padding: "20px 10px",
      display: "flex",
      flexDirection: "column",
    },
  };
  const mpModalTextFields: Partial<ITextFieldStyles> = {
    root: { margin: "10px 20px" },
    fieldGroup: {
      height: 40,
    },
  };
  const mpModalErrorTextFields: Partial<ITextFieldStyles> = {
    root: { margin: "10px 20px" },
    fieldGroup: {
      height: 40,
      border: "2px solid #F00",
    },
  };
  const mpIconStyleClass: any = mergeStyleSets({
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
        marginTop: 34.5,
        ":hover": {
          backgroundColor: "#025d60",
        },
      },
    ],
    edit: [
      {
        fontSize: 18,
        height: 14,
        width: 17,
        color: "#2392B2",
        cursor: "pointer",
      },
    ],
    delete: [
      {
        fontSize: 18,
        height: 14,
        width: 17,
        marginLeft: 10,
        color: "#CB1E06",
        cursor: "pointer",
      },
    ],
  });
  const generalStyles: any = mergeStyleSets({
    inputLabel: {
      color: "#2392B2 !important",
      display: "block",
      fontWeight: "500",
      margin: "5px 0",
      //   marginBottom: 60,
    },
    errorMessageLabel: {
      color: "#d0342c !important",
      fontWeight: 600,
      marginLeft: 20,
      fontSize: 13,
    },
  });
  // Style-Section Ends
  // States-Declaration Starts
  const [mpReRender, setMpReRender] = useState<boolean>(true);
  const [mpUnsortMasterData, setMpUnsortMasterData] =
    useState<IData[]>(mpAllItems);
  const [mpMasterData, setMpMasterData] = useState<IData[]>(mpAllItems);
  const [mpData, setMpData] = useState<IData[]>(mpAllItems);
  const [mpDisplayData, setMpDisplayData] = useState<IData[]>([]);
  const [mpFilters, setMpFilters] = useState<IFilters>(mpFilterKeys);
  const [mpcurrentPage, setMpCurrentPage] = useState<number>(currentpage);
  const [popup, setPopup] = useState<{ type: string; selectedItem: IData[] }>({
    type: "",
    selectedItem: [],
  });
  const [productValue, setProductValue] = useState<string>("");
  const [productVersionValue, setproductVersionValue] = useState<string>("");
  const [prodSubject, setProdSubject] = useState<string>("");
  const [errorStatus, setErrorStatus] = useState<string>("");
  const [mpLoader, setMpLoader] = useState<string>("noLoader");
  const [mpMasterColumns, setMpMasterColumns] = useState<IColumn[]>(mpColumns);
  // States-Declaration Ends
  //Function-Section Starts

  const versionFunction = (value) => {
    let version;
    let tempArr = mpMasterData.filter((arr) => {
      return arr.Program.toLowerCase() == (value ? value.toLowerCase() : "");
    });

    if (popup.type == "add") {
      version = "V" + (tempArr.length + 1);
    } else {
      let curProduct = mpMasterData.filter((arr) => {
        return arr.ID == popup.selectedItem[0]["ID"];
      })[0];
      if (
        curProduct["Program"].toLowerCase() !=
        (value ? value.toLowerCase() : "")
      ) {
        version = "V" + (tempArr.length + 1);
      } else {
        version = curProduct.ProductVersion;
      }
    }
    setproductVersionValue(version);
  };

  const mpGetData = (): void => {
    sharepointWeb.lists
      .getByTitle(ListName)
      .items.select("*,Author/EMail,Editor/EMail,Author/Title,Editor/Title")
      .expand("Author, Editor")
      .filter("IsDeleted ne 1")
      .orderBy("ID", true)
      .top(5000)
      .get()
      .then((items: any) => {
        items.forEach((item: any) => {
          mpAllItems.push({
            ID: item.Id ? item.Id : "",
            Program: item.Title ? item.Title : "",
            ProductVersion: item.ProductVersion ? item.ProductVersion : "",
            Subject: item.Subject ? item.Subject : "",
            Created: item.Created ? item.Created : "",
            Modified: item.Modified ? item.Modified : "",
            CreatedBy: item.Author.EMail,
            ModifiedBy: item.Editor.EMail,
            CreatedByName: item.Author.Title,
            ModifiedByName: item.Editor.Title,
          });
        });

        paginateFunction(1, mpAllItems);

        setMpUnsortMasterData([...mpAllItems]);
        columnSortArr = mpAllItems;
        setMpData([...mpAllItems]);
        columnSortMasterArr = mpAllItems;
        setMpMasterData([...mpAllItems]);
        setMpLoader("noLoader");
      })
      .catch((err) => {
        mpErrorFunction(err, "mpGetData");
      });
  };
  const mpListFilter = (key: string, option: string): void => {
    let arrBeforeFilter = [...mpMasterData];
    let tempFilterKeys = { ...mpFilters };
    tempFilterKeys[key] = option;

    if (tempFilterKeys.program) {
      arrBeforeFilter = arrBeforeFilter.filter((arr) => {
        return arr.Program.toLowerCase().includes(
          tempFilterKeys.program.toLowerCase()
        );
      });
    }

    paginateFunction(1, arrBeforeFilter);

    columnSortArr = arrBeforeFilter;
    setMpData([...columnSortArr]);
    setMpFilters({ ...tempFilterKeys });
  };
  const mpListFilterbyData = (masterData: IData[]): IData[] => {
    let arrBeforeFilter = [...masterData];
    let tempFilterKeys = { ...mpFilters };

    if (tempFilterKeys.program) {
      arrBeforeFilter = arrBeforeFilter.filter((arr) => {
        return arr.Program.toLowerCase().includes(
          tempFilterKeys.program.toLowerCase()
        );
      });
    }

    return arrBeforeFilter;
  };
  const productValidation = (value: string): any => {
    if (!value) {
      return "empty";
    }
    // else if (
    //   mpMasterData.some(
    //     (item) =>
    //       item.Program == value ||
    //       item.Program.toLowerCase() == value.toLowerCase() ||
    //       item.Program.toUpperCase() == value.toUpperCase()
    //   )
    // ) {
    //   return "exist";
    // }
    else {
      return null;
    }
  };
  const addProduct = () => {
    let masterDataBeforeUpdated = [...mpMasterData];
    sharepointWeb.lists
      .getByTitle(ListName)
      .items.add({
        Title: productValue,
        ProductVersion: productVersionValue,
        Subject: prodSubject,
      })
      .then((item) => {
        console.log(item.data);
        masterDataBeforeUpdated.push({
          ID: item.data.Id,
          Program: productValue,
          ProductVersion: productVersionValue,
          Subject: prodSubject,
          Created: item.data.Created,
          Modified: item.data.Modified,
          CreatedBy: loggeduseremail,
          ModifiedBy: loggeduseremail,
          CreatedByName: loggedusername,
          ModifiedByName: loggedusername,
        });
        paginateFunction(1, masterDataBeforeUpdated);
        setMpFilters(mpFilterKeys);
        setMpUnsortMasterData([...masterDataBeforeUpdated]);
        columnSortArr = masterDataBeforeUpdated;
        setMpData([...masterDataBeforeUpdated]);
        columnSortMasterArr = masterDataBeforeUpdated;
        setMpMasterData([...masterDataBeforeUpdated]);
        setPopup({ type: "", selectedItem: [] });
        setProductValue("");
        setproductVersionValue("");
        setErrorStatus("");
        setMpLoader("noLoader");
        ItemAddPopup();
      })
      .catch((err) => {
        mpErrorFunction(err, "addProduct");
      });
  };
  const updateProduct = (id: number) => {
    let masterDataBeforeUpdated = [...mpMasterData];
    sharepointWeb.lists
      .getByTitle(ListName)
      .items.getById(id)
      .update({
        Title: productValue,
        ProductVersion: productVersionValue,
        Subject: prodSubject,
      })
      .then((item) => {
        let targetIndex = masterDataBeforeUpdated.findIndex(
          (arr) => arr.ID == id
        );
        masterDataBeforeUpdated[targetIndex].Program = productValue;
        masterDataBeforeUpdated[targetIndex].ProductVersion =
          productVersionValue;
        masterDataBeforeUpdated[targetIndex].Subject = prodSubject;
        masterDataBeforeUpdated[targetIndex].Modified = new Date();
        masterDataBeforeUpdated[targetIndex].ModifiedBy = loggeduseremail;
        masterDataBeforeUpdated[targetIndex].ModifiedByName = loggedusername;
        let filteredArr = mpListFilterbyData(masterDataBeforeUpdated);

        paginateFunction(
          mpcurrentPage,
          filteredArr.length > 0
            ? [...filteredArr]
            : [...masterDataBeforeUpdated]
        );
        setMpFilters(
          filteredArr.length > 0 ? { ...mpFilters } : { ...mpFilterKeys }
        );
        setMpUnsortMasterData([...masterDataBeforeUpdated]);
        columnSortArr =
          filteredArr.length > 0
            ? [...filteredArr]
            : [...masterDataBeforeUpdated];
        setMpData(columnSortArr);
        columnSortMasterArr = masterDataBeforeUpdated;
        setMpMasterData([...masterDataBeforeUpdated]);
        setPopup({ type: "", selectedItem: [] });
        setProductValue("");
        setproductVersionValue("");
        setProdSubject("");
        setErrorStatus("");
        setMpLoader("noLoader");
        updatePopup();
      })
      .catch((err) => {
        mpErrorFunction(err, "updateProduct");
      });
  };
  const deleteProduct = (id: number) => {
    let masterDataBeforeUpdated = [...mpMasterData];
    // sharepointWeb.lists
    //   .getByTitle(ListName)
    //   .items.getById(id)
    //   .delete()
    //   .then((item) => {
    sharepointWeb.lists
      .getByTitle(ListName)
      .items.getById(id)
      .update({ IsDeleted: true })
      .then((item) => {
        let targetIndex = masterDataBeforeUpdated.findIndex(
          (arr) => arr.ID == id
        );

        masterDataBeforeUpdated.splice(targetIndex, 1);

        let filteredArr = mpListFilterbyData(masterDataBeforeUpdated);

        paginateFunction(
          1,
          filteredArr.length > 0
            ? [...filteredArr]
            : [...masterDataBeforeUpdated]
        );
        setMpFilters(
          filteredArr.length > 0 ? { ...mpFilters } : { ...mpFilterKeys }
        );
        setMpUnsortMasterData([...masterDataBeforeUpdated]);
        columnSortArr =
          filteredArr.length > 0
            ? [...filteredArr]
            : [...masterDataBeforeUpdated];
        setMpData(columnSortArr);
        columnSortArr = masterDataBeforeUpdated;
        setMpMasterData([...masterDataBeforeUpdated]);
        setPopup({ type: "", selectedItem: [] });
        setMpLoader("noLoader");
        deletePopup();
      })
      .catch((err) => {
        mpErrorFunction(err, "deleteProduct");
      });
  };
  const ItemAddPopup = () => (
    alertify.set("notifier", "position", "top-right"),
    alertify.success("Master product is successfully added !!!")
  );
  const updatePopup = () => (
    alertify.set("notifier", "position", "top-right"),
    alertify.success("Master product is successfully updated !!!")
  );
  const deletePopup = () => (
    alertify.set("notifier", "position", "top-right"),
    alertify.success("Master product is successfully deleted !!!")
  );
  const ErrorPopup = () => (
    alertify.set("notifier", "position", "top-right"),
    alertify.error("Something when error, please contact system admin.")
  );
  const paginateFunction = (pagenumber: number, data: IData[]) => {
    if (data.length > 0) {
      let lastIndex: number = pagenumber * totalPageItems;
      let firstIndex: number = lastIndex - totalPageItems;
      let paginatedItems = data.slice(firstIndex, lastIndex);
      currentpage = pagenumber;
      setMpDisplayData(paginatedItems);
      setMpCurrentPage(pagenumber);
    } else {
      setMpDisplayData([]);
      setMpCurrentPage(1);
    }
  };
  const _onColumnClick = (
    ev: React.MouseEvent<HTMLElement>,
    column: IColumn
  ): void => {
    const tempapColumns = mpColumns;
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
      columnSortMasterArr,
      currColumn.fieldName!,
      currColumn.isSortedDescending
    );
    setMpData([...newDisplayData]);
    setMpMasterData([...newMasterData]);
    paginateFunction(1, newDisplayData);
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
  const mpErrorFunction = (error: string, functionName: string): void => {
    console.log(error);
    let response = {
      ComponentName: "Master product list",
      FunctionName: functionName,
      ErrorMessage: JSON.stringify(error["message"]),
      Title: loggeduseremail,
    };

    Service.SPAddItem({ Listname: "Error Log", RequestJSON: response }).then(
      () => {
        setPopup({ type: "", selectedItem: [] });
        setErrorStatus("");
        setProductValue("");
        setproductVersionValue("");
        setMpLoader("noLoader");
        ErrorPopup();
      }
    );
  };

  //Function-Section Ends

  useEffect(() => {
    setMpLoader("startUpLoader");
    mpGetData();
  }, [mpReRender]);

  return (
    <>
      <div style={{ padding: "5px 10px" }}>
        {mpLoader == "startUpLoader" ? <CustomLoader /> : null}
        <div
          className={styles.mpHeaderSection}
          style={{ paddingBottom: "5px" }}
        >
          <div
            style={{ display: "flex", alignItems: "center", marginBottom: 12 }}
          >
            <Icon
              aria-label="ChevronLeftMed"
              iconName="NavigateBack"
              className={mpBigiconStyleClass.ChevronLeftMed}
              onClick={() => {
                props.handleclick("AnnualPlan");
              }}
            />
            {/* <div className={styles.mpHeader}>Master Product List</div> */}
            <Label style={{ color: "#000", fontSize: 24, padding: 0 }}>
              Product
            </Label>
            {/* <div>
              <Label style={{ marginRight: 5 }}>
                Number of records :
                <span style={{ color: "#038387" }}> {mpData.length}</span>
              </Label>
            </div> */}
          </div>
          <div>
            <button
              className={styles.mpAddBtn}
              onClick={() => {
                setPopup({ type: "add", selectedItem: [] });
              }}
            >
              Add product
            </button>
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
              justifyContent: "flex-start",
              marginBottom: 15,
            }}
          >
            <div>
              <Label styles={mpLabelStyles}>Product</Label>
              <SearchBox
                placeholder="Search product"
                styles={
                  mpFilters.program
                    ? mpActiveSearchBoxStyles
                    : mpSearchBoxStyles
                }
                value={mpFilters.program}
                onChange={(e, value: string) => {
                  mpListFilter("program", value);
                }}
              />
            </div>
            <div>
              <Icon
                iconName="Refresh"
                title="Click to reset"
                className={mpIconStyleClass.refresh}
                onClick={() => {
                  paginateFunction(1, mpUnsortMasterData);
                  columnSortArr = mpUnsortMasterData;
                  setMpData(mpUnsortMasterData);
                  columnSortMasterArr = mpUnsortMasterData;
                  setMpMasterData(mpUnsortMasterData);
                  setMpFilters({ ...mpFilterKeys });
                  setMpMasterColumns(mpColumns);
                }}
              />
            </div>
          </div>
          <div>
            <Label
              style={{
                marginTop: "30px",
                marginLeft: "10px",
                fontWeight: "500",
                color: "#323130",
                fontSize: "13px",
              }}
            >
              Number of records :
              <span style={{ color: "#038387" }}> {mpData.length}</span>
            </Label>
          </div>
        </div>
        <div>
          <DetailsList
            items={mpDisplayData}
            columns={mpMasterColumns}
            styles={mpDetailsListStyles}
            setKey="set"
            layoutMode={DetailsListLayoutMode.justified}
            selectionMode={SelectionMode.none}
          />
          {mpData.length > 0 ? (
            <div
              style={{
                display: "flex",
                justifyContent: "center",
                margin: "10px 0",
              }}
            >
              <Pagination
                currentPage={mpcurrentPage}
                totalPages={
                  mpData.length > 0
                    ? Math.ceil(mpData.length / totalPageItems)
                    : 1
                }
                onChange={(page) => {
                  paginateFunction(page, mpData);
                }}
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
        </div>
      </div>
      {popup.type == "edit" || popup.type == "add" ? (
        <Modal
          isOpen={popup.type == "edit" || popup.type == "add" ? true : false}
          isBlocking={true}
          styles={mpModalStyles}
        >
          <Label className={styles.mpPopupLabel}>
            {popup.type == "edit"
              ? "Update product"
              : popup.type == "add"
              ? "Add product"
              : ""}
          </Label>
          <div
            style={{
              display: "flex",
              marginLeft: 20,
            }}
          >
            <Label>Version : </Label>
            <Label>{productVersionValue ? productVersionValue : "V1"}</Label>
          </div>
          <div>
            <Label style={{ marginLeft: 20 }}>Product</Label>
            <TextField
              styles={errorStatus ? mpModalErrorTextFields : mpModalTextFields}
              onChange={(e, value) => {
                setErrorStatus(productValidation(value));
                versionFunction(value);
                setProductValue(value);
              }}
              placeholder="Enter a product name..."
              value={productValue}
            />
          </div>
          <div>
            <Label style={{ marginLeft: 20 }}>Subject</Label>
            <TextField
              styles={mpModalTextFields}
              onChange={(e, value) => {
                setProdSubject(value);
              }}
              placeholder="Enter subject here..."
              value={prodSubject}
            />
          </div>

          {/* <TextField
            styles={mpModalTextFields}
            disabled
            onChange={(e, value) => {}}
            placeholder="Enter a product version..."
            value={productVersionValue ? productVersionValue : "V1"}
          /> */}
          <div className={styles.mpModalBoxButtonSection}>
            {errorStatus ? (
              <Label className={generalStyles.errorMessageLabel}>
                {errorStatus == "exist"
                  ? "Product already exist"
                  : errorStatus == "empty"
                  ? "Empty value cannot be stored"
                  : null}
              </Label>
            ) : (
              ""
            )}
            <button
              className={
                errorStatus ? styles.mpSubmitBtnDisabled : styles.mpSubmitBtn
              }
              onClick={() => {
                if (productValidation(productValue) != null) {
                  setErrorStatus(productValidation(productValue));
                } else if (popup.type == "edit") {
                  setMpLoader("editLoader");
                  updateProduct(popup.selectedItem[0].ID);
                } else if (popup.type == "add") {
                  addProduct();
                  setMpLoader("addLoader");
                }
              }}
            >
              {mpLoader == "editLoader" || mpLoader == "addLoader" ? (
                <Spinner />
              ) : (
                <>
                  <Icon
                    iconName="Save"
                    style={{ position: "relative", top: 3, left: -8 }}
                  />
                  {popup.type == "edit" ? "Update" : "Submit"}
                </>
              )}
            </button>
            <button
              className={styles.mpCloseBtn}
              onClick={() => {
                setPopup({ type: "", selectedItem: [] });
                setProductValue("");
                setproductVersionValue("");
                setErrorStatus("");
                setMpLoader("noLoader");
              }}
            >
              <Icon
                iconName="Cancel"
                style={{ position: "relative", top: 3, left: -8 }}
              />
              Cancel
            </button>
          </div>
        </Modal>
      ) : popup.type == "delete" ? (
        <Modal isOpen={popup.type == "delete" ? true : false} isBlocking={true}>
          <div
            style={{
              display: "flex",
              justifyContent: "center",
              alignItems: "center",
              marginTop: "30px",
              width: "450px",
            }}
          >
            <div
              style={{
                display: "flex",
                alignItems: "center",
                justifyContent: "flex-start",
                flexDirection: "column",
                marginBottom: "10px",
              }}
            >
              <Label className={styles.deletePopupTitle}>
                Delete Master product
              </Label>
              <Label className={styles.deletePopupDesc}>
                Are you sure you want to delete the product?
              </Label>
            </div>
          </div>
          <div className={styles.mpDeletePopupBtnSection}>
            <button
              onClick={(_) => {
                deleteProduct(popup.selectedItem[0].ID);
                setMpLoader("deleteLoader");
              }}
              className={styles.mpDeletePopupYesBtn}
            >
              {mpLoader == "deleteLoader" ? <Spinner /> : "Yes"}
            </button>
            <button
              onClick={(_) => {
                setPopup({ type: "", selectedItem: [] });
                setMpLoader("noLoader");
              }}
              className={styles.mpDeletePopupNoBtn}
            >
              No
            </button>
          </div>
        </Modal>
      ) : (
        ""
      )}
    </>
  );
};

export default MasterProduct;
