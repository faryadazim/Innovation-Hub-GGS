import * as React from "react";
import { useState, useEffect } from "react";
import { sp, Web } from "@pnp/sp/presets/all";
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
  TooltipDelay,
  TooltipOverflowMode,
  DirectionalHint,
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

import Service from "../components/Services";

import alertify from "alertifyjs";
import "alertifyjs/build/css/alertify.css";

import "../ExternalRef/styleSheets/Styles.css";
import styles from "./InnovationHubIntranet.module.scss";
import Pagination from "office-ui-fabric-react-pagination";
import { mergeStyles, mergeStyleSets } from "@fluentui/react/lib/Styling";
import {
  FilePicker,
  IFilePickerResult,
} from "@pnp/spfx-controls-react/lib/FilePicker";
import CustomLoader from "./CustomLoader";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import * as Excel from "exceljs/dist/exceljs.min.js";
import * as FileSaver from "file-saver";

interface IProps {
  context: any;
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
}
interface IDropdown {
  key: string;
  text: string;
}
interface IDropdowns {
  BA: IDropdown[];
  Title: IDropdown[];
  Frequency: IDropdown[];
}
interface IConfigData {
  ID: number;
  BA: string;
  Title: string;
  Frequency: string;
  ProviderDetails: any[];
  ApproverDetails: any[];
  AudienceDetails: any[];
  Status: string;
}
interface IData {
  ID: number;
  BA: string;
  Title: string;
  Frequency: string;
  DueDate: string;
  Provider: any[];
  ApproverDetails: any[];
  AudienceDetails: any[];
  DisplayStatus: string;
  ConfigID: number;
  DocLink: string;
  TLink: string;
  Year: number;
  TimePeriod: string;
  MasterData: string;
}

let sortORData: IData[] = [];
let sortORFilterData: IData[] = [];

let CurrentPage: number = 1;
let totalPageItems: number = 10;
let DateListFormat = "DD/MM/YYYY";

const OrgMyReports = (props: IProps): JSX.Element => {
  // variable-Declaration Starts
  const sharepointWeb: any = Web(props.URL);
  const allPeoples: any[] = props.peopleList;
  const docPath = props.URL.split(".com")[1];

  const ORConfigListName = "Organisation reporting configuration list";
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

  const _ORMyReportsColumn: IColumn[] = [
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
      maxWidth: 250,
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
      maxWidth: 150,
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
                marginTop: "1px",
                cursor: "pointer",
              }}
              title={
                moment(item.DueDate, DateListFormat).format(DateListFormat) +
                ` ( ${typeAbbreviations} )`
              }
            >
              {item.DueDate
                ? moment(item.DueDate, DateListFormat).format(DateListFormat) +
                ` ( ${frequencyType} )`
                : null}
            </div>
          </>
        );
      },
    },
    {
      key: "Column6",
      name: "Status",
      fieldName: "Status",
      minWidth: 100,
      maxWidth: 200,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
      onRender: (item: IData) => (
        <>
          {item.DisplayStatus == "Read" ? (
            <div className={ORStatusStyleClass.completed}>
              {item.DisplayStatus}
            </div>
          ) : item.DisplayStatus == "Scheduled" ? (
            <div className={ORStatusStyleClass.scheduled}>
              {item.DisplayStatus}
            </div>
          ) : item.DisplayStatus == "Submitted" ? (
            <div className={ORStatusStyleClass.submitted}>
              {item.DisplayStatus}
            </div>
          ) : (
            item.DisplayStatus
          )}
        </>
      ),
    },
    {
      key: "Column7",
      name: "Audience",
      fieldName: "Audience",
      minWidth: 80,
      maxWidth: 80,
      onRender: (item) => (
        <>
          {item.AudienceDetails.length > 0 ? (
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
                  <div title={item.AudienceDetails[0].text}>
                    <Persona
                      showOverflowTooltip
                      size={PersonaSize.size24}
                      presence={PersonaPresence.none}
                      showInitialsUntilImageLoads={true}
                      imageUrl={
                        "/_layouts/15/userphoto.aspx?size=S&username=" +
                        `${item.AudienceDetails[0].secondaryText}`
                      }
                    />
                  </div>
                  {item.AudienceDetails.length > 1 ? (
                    <TooltipHost
                      content={
                        <ul style={{ margin: 10, padding: 0 }}>
                          {item.AudienceDetails.map((data, length) => {
                            return (
                              <li>
                                <div style={{ display: "flex" }}>
                                  <Persona
                                    showOverflowTooltip
                                    size={PersonaSize.size24}
                                    presence={PersonaPresence.none}
                                    showInitialsUntilImageLoads={true}
                                    imageUrl={
                                      "/_layouts/15/userphoto.aspx?size=S&username=" +
                                      `${data.secondaryText}`
                                    }
                                  />
                                  <Label style={{ marginLeft: 10 }}>
                                    {data.text}
                                  </Label>
                                </div>
                              </li>
                            );
                          })}
                        </ul>
                      }
                      delay={TooltipDelay.zero}
                      id={item.ID}
                      directionalHint={DirectionalHint.bottomCenter}
                      styles={{ root: { display: "inline-block" } }}
                    >
                      <div
                        className={styles.extraPeople}
                        aria-describedby={item.ID}
                      >
                        {item.AudienceDetails.length}
                      </div>
                    </TooltipHost>
                  ) : null}
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
      key: "Column8",
      name: "Approver",
      fieldName: "Approver",
      minWidth: 80,
      maxWidth: 80,
      onRender: (item) => (
        <>
          {item.ApproverDetails.length > 0 ? (
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
                  <div title={item.ApproverDetails[0].text}>
                    <Persona
                      showOverflowTooltip
                      size={PersonaSize.size24}
                      presence={PersonaPresence.none}
                      showInitialsUntilImageLoads={true}
                      imageUrl={
                        "/_layouts/15/userphoto.aspx?size=S&username=" +
                        `${item.ApproverDetails[0].secondaryText}`
                      }
                    />
                  </div>
                  {item.ApproverDetails.length > 1 ? (
                    <TooltipHost
                      content={
                        <ul style={{ margin: 10, padding: 0 }}>
                          {item.ApproverDetails.map((data, length) => {
                            return (
                              <li>
                                <div style={{ display: "flex" }}>
                                  <Persona
                                    showOverflowTooltip
                                    size={PersonaSize.size24}
                                    presence={PersonaPresence.none}
                                    showInitialsUntilImageLoads={true}
                                    imageUrl={
                                      "/_layouts/15/userphoto.aspx?size=S&username=" +
                                      `${data.secondaryText}`
                                    }
                                  />
                                  <Label style={{ marginLeft: 10 }}>
                                    {data.text}
                                  </Label>
                                </div>
                              </li>
                            );
                          })}
                        </ul>
                      }
                      delay={TooltipDelay.zero}
                      id={item.ID}
                      directionalHint={DirectionalHint.bottomCenter}
                      styles={{ root: { display: "inline-block" } }}
                    >
                      <div
                        className={styles.extraPeople}
                        aria-describedby={item.ID}
                      >
                        {item.ApproverDetails.length}
                      </div>
                    </TooltipHost>
                  ) : null}
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
      key: "Column9",
      name: "Upload file",
      fieldName: "Upload file",
      minWidth: 80,
      maxWidth: 80,
      onRender: (item) => (
        <>
          <div>
            <Icon
              iconName="PageArrowRight"
              title="File Upload"
              className={ORiconStyleClass.fileOpenIcon}
              onClick={() => {
                setFileUploadPopup({
                  condition: true,
                  item: item,
                  File: null,
                  fileLinkValidation: false,
                });
              }}
              styles={{
                root: {
                  fontSize: 19,
                },
              }}
            />
          </div>
        </>
      ),
    },
    {
      key: "Column10",
      name: "Open file",
      fieldName: "Doc Link",
      minWidth: 80,
      maxWidth: 80,
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
              className={
                item.DocLink
                  ? ORiconStyleClass.fileOpenIcon
                  : ORiconStyleClass.fileOpenDisabledIcon
              }
              onClick={() => { }}
            />
          </a>
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
  };
  const ORFilterOptns: IDropdowns = {
    BA: [{ key: "All", text: "All" }],
    Title: [{ key: "All", text: "All" }],
    Frequency: [{ key: "All", text: "All" }],
  };
  // variable-Declaration Starts
  // Style-Section Starts
  const ORfilterLabelStyles: Partial<ILabelStyles> = {
    root: {
      width: 150,
      marginRight: 10,
      fontSize: "13px",
      color: "#323130",
    },
  };
  const TxtBoxStyles: Partial<ITextFieldStyles> = {
    root: {
      width: 300,
    },
    field: {
      fontSize: 12,
      color: "#000",
      borderRadius: 4,
      background: "#fff !important",
    },
    fieldGroup: {
      border: "1px solid #000 !important",
      height: "36px",
    },
  };
  const ErrorTxtBoxStyles: Partial<ITextFieldStyles> = {
    root: {
      width: 300,
    },
    field: {
      fontSize: 12,
      color: "#000",
      borderRadius: 4,
      background: "#fff !important",
    },
    fieldGroup: {
      border: "2px solid #f00 !important",
      height: "36px",
    },
  };
  const ORStatusStyle = mergeStyles({
    textAlign: "center",
    borderRadius: "25px",
    width: "160px",
  });
  const ORStatusStyleClass = mergeStyleSets({
    completed: [
      {
        fontWeight: "600",
        padding: "3px",
        color: "#187B29",
        backgroundColor: "#D4FFDB",
      },
      ORStatusStyle,
    ],
    scheduled: [
      {
        fontWeight: "600",
        padding: "3px",
        color: "#06637E",
        backgroundColor: "#97E9EC",
      },
      ORStatusStyle,
    ],
    submitted: [
      {
        fontWeight: "600",
        padding: "3px",
        color: "#895C09 ",
        backgroundColor: "#FFDB99",
      },
      ORStatusStyle,
    ],
    overdue: [
      {
        fontWeight: "600",
        padding: "3px",
        color: "#CB1E06",
        backgroundColor: "#FFD3CD",
      },
      ORStatusStyle,
    ],
    pending: [
      {
        fontWeight: "600",
        padding: "3px",
        color: "#773030",
        backgroundColor: "#e6b1b1",
      },
      ORStatusStyle,
    ],
  });
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
  const ORFileUploadDescStyles = mergeStyleSets({
    heading: {
      color: "#323130",
      fontSize: 16,
      marginLeft: 10,
      fontWeight: 600,
    },
    DescHeadingLeft: { color: "#000", fontSize: 16, fontWeight: 500 },
    DescHeadingRight: {
      color: "#2392B2",
      fontSize: 16,
      fontWeight: 500,
      marginLeft: 5,
    },
  });
  const ORiconStyleClass = mergeStyleSets({
    refresh: {
      color: "white",
      fontSize: "18px",
      height: 22,
      width: 22,
      cursor: "pointer",
      backgroundColor: "#038387",
      padding: 5,
      marginTop: 22,
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
      userSelect: "none",
      color: "#038387",
      fontSize: 22,
      height: 22,
      width: 22,
      cursor: "pointer",
      marginRight: 5,
      marginLeft: "25px",
    },
    fileOpenDisabledIcon: {
      userSelect: "none",
      color: "#ababab",
      fontSize: 22,
      height: 22,
      width: 22,
      cursor: "not-allowed",
      marginRight: 5,
      marginLeft: "25px",
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
  const [ORColumns, setORColumns] = useState<IColumn[]>(_ORMyReportsColumn);
  const [fileUploadPopup, setFileUploadPopup] = useState<{
    condition: boolean;
    item: IData;
    File: any;
    fileLinkValidation: boolean;
  }>({ condition: false, item: null, File: null, fileLinkValidation: false });
  const [ORCurrentPage, setORCurrentPage] = useState<number>(CurrentPage);
  const [ORLoader, setORLoader] = useState("noLoader");
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
      { header: "Audience", key: "Audience", width: 25 },
      { header: "Approver", key: "Approver", width: 25 },
      { header: "Status", key: "Status", width: 60 },
    ];
    arrExport.forEach((item) => {
      let Audience = "";
      item.AudienceDetails.length > 0
        ? item.AudienceDetails.forEach((arr) => {
          Audience += arr.text + ";";
        })
        : null;

      let Approver = "";
      item.ApproverDetails.length > 0
        ? item.ApproverDetails.forEach((arr) => {
          Approver += arr.text + ";";
        })
        : null;

      worksheet.addRow({
        BA: item.BA ? item.BA : "",
        Title: item.Title ? item.Title : "",
        Frequency: item.Frequency ? item.Frequency : "",
        Status: item.DisplayStatus ? item.DisplayStatus : "",
        DueDate: item.DueDate
          ? moment(item.DueDate, DateListFormat).format(DateListFormat)
          : "",
        Audience: Audience ? Audience : "",
        Approver: Approver ? Approver : "",
      });
    });
    ["A1", "B1", "C1", "D1", "E1", "F1", "G1"].map((key) => {
      worksheet.getCell(key).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "00e8d1" },
      };
    });
    ["A1", "B1", "C1", "D1", "E1", "F1", "G1"].map((key) => {
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
          `Organisationreport-Myreport-${new Date().toLocaleString()}.xlsx`
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
      ComponentName: "Org reporting - myReports",
      FunctionName: functionName,
      ErrorMessage: JSON.stringify(error["message"]),
      Title: currentLoggedUserEmail,
    };

    Service.SPAddItem({ Listname: "Error Log", RequestJSON: response }).then(
      () => {
        setFileUploadPopup({
          condition: false,
          item: null,
          File: null,
          fileLinkValidation: false,
        });
        setORLoader("noLoader");
        ErrorPopup();
      }
    );
  };
  const fileUploadValiadtion = (): void => {
    if (fileUploadPopup.File == null) {
      fileUploadPopup.fileLinkValidation = true;
    }

    if (fileUploadPopup.fileLinkValidation) {
      setORLoader("noLoader");
      setFileUploadPopup({ ...fileUploadPopup });
    } else {
      setFileUploadPopup({ ...fileUploadPopup });
      fileUploadFunction();
    }
  };

  const reloadFilterDropdowns = (data: IData[]): void => {
    let tempArrReload = data;

    tempArrReload.forEach((item) => {
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
    });

    setORFilterDrpDown(ORFilterOptns);
  };

  // const fileUploadFunction = (): void => {
  //   fileUploadPopup.File[0]
  //     .downloadFileContent()
  //     .then((r) => {
  //       sharepointWeb
  //         .getFolderByServerRelativeUrl(`${docPath}` + "/OrgReportingDocuments")
  //         .files.add(fileUploadPopup.File[0].fileName, r, true)
  //         .then((item: any) => {
  //           item.file.getItem().then((fileItem: any) => {
  //             fileItem
  //               .update({
  //                 ConfigID: fileUploadPopup.item.ConfigID,
  //                 OrgReportID: fileUploadPopup.item.ID,
  //                 Frequency: fileUploadPopup.item.Frequency,
  //                 Year: fileUploadPopup.item.Year,
  //                 TimePeriod: fileUploadPopup.item.TimePeriod,
  //               })
  //               .then((updatedItem: any) => {
  //                 sharepointWeb.lists
  //                   .getByTitle(OrgReportListName)
  //                   .items.getById(fileUploadPopup.item.ID)
  //                   .get()
  //                   .then((_item) => {
  //                     sharepointWeb.lists
  //                       .getByTitle(OrgReportListName)
  //                       .items.getById(fileUploadPopup.item.ID)
  //                       .update({
  //                         DisplayStatus: "Submitted",
  //                         DocLink: `${props.URL}/OrgReportingDocuments/${fileUploadPopup.File[0].fileName}`,
  //                         ApproverActioned: 0,
  //                         Comments: "",
  //                         Status: _item.Status ? "Not actioned" : null,
  //                         SubmittedOn: moment().format("YYYY/MM/DD"),
  //                       })
  //                       .then((_r: any) => {
  //                         fileLinkAutoPopulateFunction(
  //                           fileUploadPopup.item,
  //                           `${props.URL}/OrgReportingDocuments/${fileUploadPopup.File[0].fileName}`
  //                         );
  //                       })
  //                       .catch((error: any) => {
  //                         ORErrorFunction(error, "orgReportingUpdateFunction");
  //                       });
  //                   })
  //                   .catch((error: any) => {
  //                     ORErrorFunction(error, "getProviderData");
  //                   });
  //               })
  //               .catch((error: any) => {
  //                 ORErrorFunction(error, "filecolumnUploadFunction");
  //               });
  //           });
  //         })
  //         .catch((error) => {
  //           ORErrorFunction(error, "fileUploadFunction");
  //         });
  //     })
  //     .catch(ORErrorFunction);
  // };

  const fileUploadFunction = (): void => {
    sharepointWeb.lists
      .getByTitle(OrgReportListName)
      .items.getById(fileUploadPopup.item.ID)
      .get()
      .then((_item) => {
        sharepointWeb.lists
          .getByTitle(OrgReportListName)
          .items.getById(fileUploadPopup.item.ID)
          .update({
            DisplayStatus: "Submitted",
            DocLink: fileUploadPopup.File,
            ApproverActioned: 0,
            Comments: "",
            Status: _item.Status ? "Not actioned" : null,
            SubmittedOn: moment().format("YYYY/MM/DD"),
          })
          .then((_r: any) => {
            fileLinkAutoPopulateFunction(
              fileUploadPopup.item,
              fileUploadPopup.File
            );
          })
          .catch((error: any) => {
            ORErrorFunction(error, "orgReportingUpdateFunction");
          });
      })
      .catch((error: any) => {
        ORErrorFunction(error, "getProviderData");
      });
  };
  const fileLinkAutoPopulateFunction = (
    _item: IData,
    _DocLink: string
  ): void => {
    let count: number = 0;
    sharepointWeb.lists
      .getByTitle(OrgReportListName)
      .items.filter(`ParentID eq ${_item.ID}`)
      .top(5000)
      .get()
      .then((_responseData: any) => {
        if (_responseData.length > 0) {
          _responseData.forEach((_item) => {
            sharepointWeb.lists
              .getByTitle(OrgReportListName)
              .items.getById(_item.ID)
              .update({
                DocLink: _DocLink,
                Status: "Not actioned",
                Comments: "",
                SubmittedOn: moment().format("YYYY/MM/DD"),
                ActionedOn: null,
              })
              .then((_r: any) => {
                count++;
                if (_responseData.length == count) {
                  setFileUploadPopup({
                    condition: false,
                    item: null,
                    File: null,
                    fileLinkValidation: false,
                  });
                  setORLoader("noLoader");
                  setORReRender(!ORReRender);
                }
              })
              .catch((err) => {
                ORErrorFunction(err, "fileLinkAutoPopulateFunctionLoop");
              });
          });
        } else {
          setFileUploadPopup({
            condition: false,
            item: null,
            File: null,
            fileLinkValidation: false,
          });
          setORLoader("noLoader");
          setORReRender(!ORReRender);
        }
      })
      .catch((err) => {
        ORErrorFunction(err, "fileLinkAutoPopulateFunction");
      });
  };

  // getData function
  const getConfigData = (): void => {
    let _ORConfigData: IConfigData[] = [];
    sharepointWeb.lists
      .getByTitle(ORConfigListName)
      .items.filter("Inactive ne 1")
      .top(5000)
      .orderBy("Modified", false)
      .get()
      .then((items) => {
        items.forEach((item) => {
          let _ApproverDetails = [];
          let _AudienceDetails = [];
          let _ProviderDetails = [];
          if (item.ApproverId != null && item.ApproverId.length > 0) {
            item.ApproverId.forEach((user) => {
              _ApproverDetails.push(
                allPeoples.filter((people) => {
                  return people.ID == user;
                })[0]
              );
            });
          }
          if (item.AudienceId != null && item.AudienceId.length > 0) {
            item.AudienceId.forEach((user) => {
              _AudienceDetails.push(
                allPeoples.filter((people) => {
                  return people.ID == user;
                })[0]
              );
            });
          }
          if (item.ProviderId != null && item.ProviderId) {
            _ProviderDetails.push(
              allPeoples.filter((people) => {
                return people.ID == item.ProviderId;
              })[0]
            );
          }
          _ORConfigData.push({
            ID: item.ID,
            BA: item.BA,
            Title: item.Title,
            Frequency: item.Frequency,
            ApproverDetails: [..._ApproverDetails],
            AudienceDetails: [..._AudienceDetails],
            ProviderDetails: [..._ProviderDetails],
            Status: item.Status,
          });
        });
        getOrgReportData(_ORConfigData);
      })
      .catch((error) => {
        ORErrorFunction(error, "getOrgReportConfig");
      });
  };
  const getOrgReportData = (configData: IConfigData[]): void => {
    let _ORdata: IData[] = [];
    sharepointWeb.lists
      .getByTitle(OrgReportListName)
      .items.select("*", "FieldValuesAsText/DueDate")
      .expand("FieldValuesAsText")
      // .filter(
      //   `UserId eq ${currentLoggedUserID} and Year eq ${OR_Year} and GroupID eq 1 and MasterData eq 'Yes' and Inactive ne 1`
      // )
      .orderBy("ID", false)
      .top(5000)
      .get()
      .then((items: any[]) => {
        items.forEach((item) => {

          if (item.UserId == currentLoggedUserID && item.Year == OR_Year && item.GroupID == 1  ) {
            let _ApproverDetails = configData.filter((_item: IConfigData) => {
              return _item.ID == item.ConfigID;
            })[0]?.ApproverDetails;
            let _AudienceDetails = configData.filter((_item: IConfigData) => {
              return _item.ID == item.ConfigID;
            })[0]?.AudienceDetails;
            let _ProviderDetails = configData.filter((_item: IConfigData) => {
              return _item.ID == item.ConfigID;
            })[0]?.ProviderDetails;
            _ORdata.push({
              ID: item.ID,
              BA: item.BA,
              Title: item.Title,
              Frequency: item.Frequency,
              DueDate: moment(
                item["FieldValuesAsText"].DueDate,
                DateListFormat
              ).format(DateListFormat),
              Provider: [..._ProviderDetails],
              ApproverDetails: [..._ApproverDetails],
              AudienceDetails: [..._AudienceDetails],
              ConfigID: item.ConfigID,
              DocLink: item.DocLink,
              TLink: item.TLink,
              Year: item.Year,
              TimePeriod: item.TimePeriod,
              DisplayStatus: item.DisplayStatus,
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
    const tempORColumns = _ORMyReportsColumn;
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
    getConfigData();
  }, [ORReRender]);
  return (
    <div style={{ marginTop: "5px" }}>
      {ORLoader == "StartLoader" ? (
        <CustomLoader />
      ) : (
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
                    setORColumns(_ORMyReportsColumn);
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
          <>
            {fileUploadPopup.condition ? (
              <Modal isOpen={fileUploadPopup.condition} isBlocking={false}>
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
                    File upload
                  </div>
                  <div style={{ marginBottom: 10 }}>
                    <div
                      style={{
                        display: "flex",
                        justifyContent: "space-between",
                      }}
                    >
                      <div style={{ display: "flex" }}>
                        <Label
                          className={ORFileUploadDescStyles.DescHeadingLeft}
                        >
                          Title :
                        </Label>
                        <Label
                          className={ORFileUploadDescStyles.DescHeadingRight}
                        >
                          {fileUploadPopup.item.Title}
                        </Label>
                      </div>

                      <div style={{ display: "flex" }}>
                        <Label
                          className={ORFileUploadDescStyles.DescHeadingLeft}
                        >
                          Frequency :
                        </Label>
                        <Label
                          className={ORFileUploadDescStyles.DescHeadingRight}
                        >
                          {fileUploadPopup.item.Frequency}
                        </Label>
                      </div>
                    </div>
                  </div>
                  <div>
                    <Label required={true} style={{ width: 500 }}>
                      Link
                    </Label>
                    <TextField
                      placeholder="Add link"
                      value={fileUploadPopup.File}
                      styles={
                        fileUploadPopup.fileLinkValidation
                          ? ErrorTxtBoxStyles
                          : TxtBoxStyles
                      }
                      onChange={(e, value: string) => {
                        let tempFileUpload = { ...fileUploadPopup };
                        tempFileUpload.File = value;
                        tempFileUpload.fileLinkValidation = false;
                        setFileUploadPopup({ ...tempFileUpload });
                      }}
                    />
                  </div>
                  {/* <div
                    style={{
                      width: 500,
                      display: "flex",
                      justifyContent: "space-between",
                    }}
                  >
                    <div
                      style={{
                        display: "flex",
                      }}
                    >
                      <div className={styles.filebtnWrapper}>
                        <FilePicker
                          buttonClassName={styles.filebtn}
                          bingAPIKey="<BING API KEY>"
                          buttonLabel="Upload file"
                          accepts={[
                            ".gif",
                            ".jpg",
                            ".jpeg",
                            ".bmp",
                            ".dib",
                            ".tif",
                            ".tiff",
                            ".ico",
                            ".png",
                            ".jxr",
                            ".svg",
                            ".pdf",
                            ".xls",
                            ".xlsx",
                            ".doc",
                            ".docx",
                            ".ppt",
                            ".pptx",
                            ".pdf",
                          ]}
                          buttonIcon="FileImage"
                          onSave={(filePickerResult: IFilePickerResult[]) => {
                            let tempFileUpload = { ...fileUploadPopup };
                            filePickerResult.length > 0
                              ? (tempFileUpload.File = filePickerResult)
                              : (tempFileUpload.File = []);
                            setFileUploadPopup({ ...tempFileUpload });
                          }}
                          context={props.spcontext}
                        />
                      </div>
                      <div>
                        {fileUploadPopup.File != null &&
                        fileUploadPopup.File.length > 0 ? (
                          <Label>{fileUploadPopup.File[0].fileName}</Label>
                        ) : null}
                      </div>
                    </div>
                    <div></div>
                  </div> */}
                  <div className={styles.ORFileUploadModalBoxButtonSection}>
                    {fileUploadPopup.fileLinkValidation ? (
                      <Label style={{ color: "#f00", fontWeight: 600 }}>
                        * All fields are mandatory
                      </Label>
                    ) : null}
                    <button
                      className={styles.ORModalBoxSubmitBtn}
                      onClick={(_) => {
                        if (ORLoader == "noLoader") {
                          setORLoader("onModalSubmit");
                          fileUploadValiadtion();
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
                          {"Upload"}
                        </span>
                      )}
                    </button>
                    <button
                      className={styles.ORModalBoxBackBtn}
                      onClick={(_) => {
                        if (ORLoader == "noLoader") {
                          setFileUploadPopup({
                            condition: false,
                            item: null,
                            File: null,
                            fileLinkValidation: false,
                          });
                        }
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
          </>
          {/* Modal-Section Ends */}
        </div>
      )}
    </div>
  );
};

export default OrgMyReports;
