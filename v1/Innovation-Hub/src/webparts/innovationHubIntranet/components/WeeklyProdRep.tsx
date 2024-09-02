import * as React from "react";
import { useState, useEffect } from "react";
import * as moment from "moment";
import { Web } from "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import * as Excel from "exceljs/dist/exceljs.min.js";
import * as FileSaver from "file-saver";
import './customStyle.css'

import {
    Icon, Label, Dropdown, TooltipHost, IDropdownStyles,ILabelStyles
} from "@fluentui/react";
import "react-quill/dist/quill.snow.css";
import "../ExternalRef/styleSheets/Styles.css";
import { mergeStyles, mergeStyleSets } from "@fluentui/react/lib/Styling";
import styles from "./InnovationHubIntranet.module.scss";
import CustomLoader from "./CustomLoader";
import alertify from "alertifyjs";
import "alertifyjs/build/css/alertify.css";

import { IStackTokens } from '@fluentui/react/lib/Stack';
import { Toggle } from '@fluentui/react/lib/Toggle';

import {
    DatePicker,
    DayOfWeek,
    defaultDatePickerStrings,
} from '@fluentui/react';


import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { IStackProps, Stack } from '@fluentui/react/lib/Stack';


const rootClass = mergeStyles({ maxWidth: 300, selectors: { '> *': { marginBottom: 15 } } });
const firstDayOfWeekForDatePicker = DayOfWeek.Monday;
const PLIconStyle = mergeStyles({
    fontSize: 17,
    height: 14,
    width: 17,
    cursor: "pointer",
});
const drDropdownStyles: Partial<IDropdownStyles> = {
    root: {
        width: 145,
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
    callout: {
        maxHeight: "400px !important",
    },
    caretDown: { fontSize: 14, color: "#000" },
};
const drActiveDropdownStyles: Partial<IDropdownStyles> = {
    root: {
        width: 145,
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
    callout: {
        maxHeight: "400px !important",
    },
    caretDown: { fontSize: 14, color: "#038387" },
};
const drLabelStyles: Partial<ILabelStyles> = {
    root: {
      width: 150,
      marginRight: 10,
      fontSize: "13px",
      color: "#323130",
    },
  };
  const drToggleLabelStyles: Partial<ILabelStyles> = {
    root: {
      width: 94,
      marginRight: 10,
      fontSize: "13px",
      color: "#323130",
    },
  };
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


const rowProps: IStackProps = { horizontal: true, verticalAlign: 'center' };

const stackTokens: IStackTokens = { childrenGap: 10 };
const tokens = {
    sectionStack: {
        childrenGap: 10,
    },
    spinnerStack: {
        childrenGap: 20,
    },
};
function convertDateFormat(input: string): string {
    // Parse input to a Date object
    const date = new Date(input);

    // Extract date, month, and year
    const day = date.getUTCDate();
    const month = date.getUTCMonth() + 1; // Months are 0-indexed in JavaScript
    const year = date.getUTCFullYear();

    // Pad day and month with leading 0 if necessary
    const paddedDay = day < 10 ? '0' + day : day;
    const paddedMonth = month < 10 ? '0' + month : month;

    // Format and return the date in the desired format
    return `${paddedDay}/${paddedMonth}/${year}`;
}




const BusinessArea = (props: any) => {
    let years = ["2022", "2023", "2024"];
    let weeks = [
        'Week 1', 'Week 2', 'Week 3', 'Week 4',
        'Week 5', 'Week 6', 'Week 7', 'Week 8',
        'Week 9', 'Week 10', 'Week 11', 'Week 12',
        'Week 13', 'Week 14', 'Week 15', 'Week 16',
        'Week 17', 'Week 18', 'Week 19', 'Week 20',
        'Week 21', 'Week 22', 'Week 23', 'Week 24',
        'Week 25', 'Week 26', 'Week 27', 'Week 28',
        'Week 29', 'Week 30', 'Week 31', 'Week 32',
        'Week 33', 'Week 34', 'Week 35', 'Week 36',
        'Week 37', 'Week 38', 'Week 39', 'Week 40',
        'Week 41', 'Week 42', 'Week 43', 'Week 44',
        'Week 45', 'Week 46', 'Week 47', 'Week 48',
        'Week 49', 'Week 50', 'Week 51', 'Week 52',
        'Week 53'];
    let RevListName = "Review Log"
    // Variable-Declaration-Section Starts
    let main_url = "https://ggsaus.sharepoint.com";
    const sharepointWeb = Web(main_url);
    let DateListFormat = "YYYY-MM-DD";

    // Styles-Section Ends
    // States-Declaration Starts
    const [drReRender, setDrReRender] = useState(true);
    const [firstDate, setFirstDate] = useState(moment('2021-12-12', 'YYYY-MM-DD').format('YYYY-MM-DD'));
    const [lastDate, setLastDate] = useState(moment('2024-12-12', 'YYYY-MM-DD').format('YYYY-MM-DD'));
    const [DisplayData, setSlDisplayData] = useState([]);
    let [selectedYear, setSelectedYear] = useState("");
    let [selectedWeek, setSelectedWeek] = useState("");
    const [drLoader, setDrLoader] = useState("noLoader");
    const [businessWiseData, setBusinessWiseData] = useState([])
    const [yearList, setYearList] = useState([...years])
    const [weekList, setWeekList] = useState([...weeks])
    const [isAvailableData, setIsAvailableData] = useState(false)
    const [isSpinnerOn, setIsSpinnerOn] = useState(false)
    const [isToggleOn, setIsToggleOn] = useState(true)

    // const [fromDate, setfromDate] = useState(moment('2021-12-12', 'YYYY-MM-DD').format('YYYY-MM-DD'));
    // const [toDate, setToDate] =useState(moment('2024-12-12', 'YYYY-MM-DD').format('YYYY-MM-DD'));

    const _onChangeToggle = (ev: any, checked?: any) => {
        setIsToggleOn(checked)
    }

    function generateWeeksList(year) {
        const weeksInYear = moment(year, 'YYYY').isoWeeksInYear();
        const weeksList_ = [];
        for (let i = 1; i <= weeksInYear; i++) {
            weeksList_.push(`Week ${i}`);
        }
        setWeekList(weeksList_);
        setSelectedWeek(weeksList_[0])
    }

    const drGetAllOptions = async () => {
        setIsSpinnerOn(true)
        let userList = [];
        let aplistData = [];

        await sharepointWeb.lists
            .getByTitle("Annual Plan")
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
                "Master_x0020_Project/ProductVersion",
                "FieldValuesAsText/StartDate",
                "FieldValuesAsText/PlannedEndDate"
            )
            .expand(
                "ProjectOwner",
                "ProjectLead",
                "Master_x0020_Project",
                "FieldValuesAsText"
            )
            .top(5000)
            .orderBy("Modified", false)
            .get()
            .then((items) => {
                //apply date condition here
                //abstracting user from annual plan  

                console.log(items)
                items.forEach((item: any, index: number) => {


                    if (item.StartDate) {

                        const eventStartDate = moment(item.StartDate.slice(0, 10), 'YYYY-MM-DD');
                        const firstDate_ = moment(firstDate, "YYYY-MM-DD");
                        const lastDate_ = moment(lastDate, "YYYY-MM-DD");

                        if (
                            (eventStartDate.isBetween(firstDate_, lastDate_, undefined, '[]')) &&
                            (item.Status == "On hold" || item.Status == "Scheduled"
                                || item.Status == "Unplanned work " || item.Status == "Overdue" || item.Status == "Overdue" || item.Status == "Behind schedule")
                        ) {

                            aplistData.push({
                                ...item,
                                diliverable_ph: 0,
                                diliverable_ah: 0,
                                taskLists: [],
                                playBooks: [],
                            })



                            // fetching user list 
                            if (item.ProjectLeadId && item.ProjectLeadId != null && item.ProjectLeadId !== null) {
                                item.ProjectLead.map((x) => {
                                    if (userList.length > 0 && userList.findIndex(y => y.Id == x.Id) != -1) {

                                    } else {



                                        userList.push({
                                            Title: x.Title,
                                            Id: x.Id,
                                            Email: x.EMail,
                                            Role: "Role Assign",
                                            deliverable_list: [],
                                            logs_list: []
                                        })
                                    }
                                })
                            }
                        }
                    }

                });
            })
            .catch((err) => {
                console.log(err), "error";
            });


        // console.log(aplistData, "aplistData")

        let annualPlanItemWithTaskLists = await getPbData(aplistData)
        let countObj = {};
        // Initialize all business areas with 0 values
        annualPlanItemWithTaskLists.forEach((item: any) => {
            const businessArea = item?.BusinessArea;
            if (!countObj.hasOwnProperty(businessArea)) {
                countObj[businessArea] = {
                    PlannedHours: 0,
                    ActualHours: 0,
                };
            }
        });

        //fetching all review logs 
        let item_rev_log = []
        await sharepointWeb.lists
            .getByTitle(RevListName)
            .items.select(
                "*",
                "FromUser/Title",
                "FromUser/Id",
                "FromUser/EMail",
                "CcEmail/Title",
                "CcEmail/Id",
                "CcEmail/EMail"
            )
            .expand("FromUser,CcEmail")
            .top(5000)
            .orderBy("Modified", false)
            .get()
            .then(async (items_rev_log) => {
                // item_rev_log = items_rev_log;


                // if (item.StartDate) {

                //     const eventStartDate = moment(item.StartDate.slice(0, 10), 'YYYY-MM-DD');
                //     const firstDate_ = moment(firstDate, "YYYY-MM-DD");
                //     const lastDate_ = moment(lastDate, "YYYY-MM-DD");

                //     if (
                //         (eventStartDate.isBetween(firstDate_, lastDate_, undefined, '[]')) &&



                items_rev_log.forEach(item => {
                    if (item.auditSent) {


                        // const eventStartDate = moment(item.auditSent.slice(0, 10), 'YYYY-MM-DD');
                        //     const firstDate_ = moment(firstDate, "YYYY-MM-DD");
                        //     const lastDate_ = moment(lastDate, "YYYY-MM-DD");
                        //     if      (eventStartDate.isBetween(firstDate_, lastDate_, undefined, '[]')){

                        item_rev_log.push(item)
                        // }


                    }


                })


            })
            .catch((err) => {
                console.log(err, "drReallocateFunction-getItem");
            });

        let count = 0

        annualPlanItemWithTaskLists.forEach((item: any, index: number) => {

            if (item.ProjectLeadId && item.ProjectLeadId != null && item.ProjectLeadId !== null) {
                item.ProjectLead.map((l) => {
                    let proj_lead_index = userList.findIndex(w => w.Id == l.Id)
                    // find specffic user and add adding annualplain list to specif user 
                    const deliverable_list = userList[proj_lead_index].deliverable_list;
                    // console.log(deliverable_list, "INSIDE ASSINGING LOOP")

                    const businessArea = item?.BusinessArea;
                    const tasklists = item?.taskLists;

                    //code to calcuate oh/ah bussiness vise from all anmilies 
                    tasklists.forEach(task => {
                        const PlannedHours = task.PlannedHours;
                        const ActualHours = task.ActualHours;
                        if (countObj.hasOwnProperty(businessArea)) {
                            countObj[businessArea].PlannedHours += PlannedHours;
                            countObj[businessArea].ActualHours += ActualHours;
                        } else {
                            countObj[businessArea] = {
                                PlannedHours: PlannedHours,
                                ActualHours: ActualHours
                            };
                        }
                    });

                    const project_specific_docs: any[] = item_rev_log.filter((log_item) => log_item?.AnnualPlanID == item?.Id)
                    const list_test = [];
                    project_specific_docs.map((er) => list_test.push({
                        anomilies: er?.Title,
                        request: er?.auditRequestType,
                        request_date: er?.auditSent,
                        sent_to: er?.auditTo,
                        response: er.auditResponseType,
                        response_date: er?.auditResponseDate,
                    }))
                    // console.log(list_test, "Un Sorted")
                    const sortedFilteredAnnomilies = annomiliesformator(list_test)
                    //  console.log(sortedFilteredAnnomilies, "Sorted")

                    // console.log(item?.Title, " item?.Title")

                    count++;
                    userList[proj_lead_index] = {
                        ...userList[proj_lead_index],
                        deliverable_list: [...deliverable_list, {
                            client: item?.ProjectOwner?.Title,
                            date: item?.StartDate,
                            status: item?.Status,
                            terms: item?.TermNew,
                            deliverable_name: item?.Title,
                            diliverable_ph: item?.diliverable_ph,
                            diliverable_ah: item?.diliverable_ah,
                            BA: item?.BusinessArea,

                            taskLists: item?.taskLists,
                            playbooks: item?.playBooks,
                            documents_review_log: [...sortedFilteredAnnomilies]
                        }]
                    }
                })
            }
        });

        // console.log(userList, "userList")
        const dataArray = Object.keys(countObj).map(key => ({
            businessArea: key,
            plannedHours: countObj[key].PlannedHours,
            actualHours: countObj[key].ActualHours,
        }));
        dataArray.sort((a, b) => a.businessArea.localeCompare(b.businessArea));

        setBusinessWiseData(dataArray);
        //   divinding user according to annual plan 
        filteredByBusinessArea(userList)
        setIsAvailableData(true)
        // setSlDisplayData(userList);
        setDrLoader("noLoader");

    };

    function checkIfEndsWithArchive(str) {
        return str.toLowerCase().endsWith("archive");
    }
    const YearFinder = async () => {
        // ...................
        let year_ = [];
        await sharepointWeb.lists
            .getByTitle("Annual Plan")
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
                "Master_x0020_Project/ProductVersion",
                "FieldValuesAsText/StartDate",
                "FieldValuesAsText/PlannedEndDate"
            )
            .expand(
                "ProjectOwner",
                "ProjectLead",
                "Master_x0020_Project",
                "FieldValuesAsText"
            )
            .top(5000)
            .orderBy("Modified", false)
            .get()
            .then((items) => {
                //apply date condition here
                //abstracting user from annual plan 

                items.forEach((item: any, index: number) => {

                    year_.push(`${item.Year}`);
                });

            })
            .catch((err) => {
                console.log(err), "error";
            });

        var uniqueYears_ = year_.filter(function (value, index, self) {
            return value !== 'null' && self.indexOf(value) === index;
        });

        if (uniqueYears_.length < 1) {
            // show here toast that no  data found in annual plan 
            // console.log("Not Data Found")
        } else {
            // console.log("Data Found", uniqueYears_)
            setYearList(uniqueYears_)
            setSelectedYear(uniqueYears_[0]);
            generateWeeksList(uniqueYears_[0])
            setDrLoader("noLoader");
        }

        //select first value as sleected year
        //fetch week of year and feetch week for all
    };

    const annomiliesformator = (data) => {
        //here this code first pick a anomlies and then add its 
        const uniqueResult = data.reduce((acc, item) => {
            if (!acc[item.anomilies]) {
                acc[item.anomilies] = {
                    anomilies: item.anomilies,
                    request: [],
                    response: [],
                    sent_to: [],
                    request_date: item?.request_date,
                    response_date: item?.response_date,
                };
            }

            acc[item.anomilies].request.push(item.request);
            acc[item.anomilies].response.push(item.response);
            acc[item.anomilies].sent_to.push(item.sent_to);

            return acc;
        }, {});

        var outputArray = Object.keys(uniqueResult).map(function (key) {
            var value = uniqueResult[key];
            // if(value?.response?.length>1 

            return {
                anomilies: value?.anomilies,
                request: value?.request,
                response: value?.response,
                sent_to: value?.sent_to,
                request_date: value?.request_date,
                response_date: value?.response_date,
            };

        });

        const filteredDocuments = outputArray.filter((x) => {
            // currentl we are fetching us time make it australia if need **
            const currentDateTime = moment().utc().format('YYYY-MM-DDTHH:mm:ss[Z]');

            let req_date = currentDateTime;
            let res_date = currentDateTime;

            if (x.request_date !== null) {
                req_date = x.request_date;
            }
            if (x.response_date !== null) {
                res_date = x.response_date;
            }

            const req_date_moment = moment(req_date);
            const res_date_moment = moment(res_date);
            const diff = moment.duration(res_date_moment.diff(req_date_moment)).asHours();
            return x?.response?.length > 1 && diff > 48
        })
        //  filteredDocuments;



        let eleboratedFilteredDocuments = [];

        filteredDocuments.forEach(obj => {
            let len = obj.request.length;
            for (let i = 0; i < len; i++) {
                eleboratedFilteredDocuments.push({
                    anomilies: obj.anomilies,
                    request: [obj.request[i]],
                    request_date: obj.request_date,
                    response: [obj.response[i]],
                    response_date: obj.response_date,
                    sent_to: [obj.sent_to[i]]
                });
            }
        });

        // console.log(eleboratedFilteredDocuments);
        return eleboratedFilteredDocuments



    }
    const filteredByBusinessArea = (userList) => {

        const result: any[] = [];
        const businessAreas: any = {};

        userList.forEach((item) => {
            item.deliverable_list.forEach((delivery) => {


                const { BA } = delivery;
                //agr bussiness area hy?
                if (businessAreas[BA]) {
                    const existingUser = businessAreas[BA]?.user_info?.some((user) => user?.userinfo.Id === item?.Id);
                    //although im using here some function but it will be always one user in each bussiness area
                    if (existingUser) {
                        //agr us bussiness area mein vo user hy?
                        // let prevRec = existingUser?.diliveralbe_list;
                        existingUser?.diliveralbe_list?.push(delivery);
                        //inf bussiness area inisde user and store this dilivery in that user 
                        let userIndexInsideBA = businessAreas[BA]?.user_info.findIndex(o => o?.userinfo.Id == item?.Id);
                        if (userIndexInsideBA > -1) {
                            // console.log(userIndexInsideBA, "exist")
                            businessAreas[BA]?.user_info[userIndexInsideBA].diliveralbe_list.push(delivery);
                        } else {

                            // console.log(userIndexInsideBA, "not exist")
                        }

                        //im not here updating bussiness area anymore
                    } else {
                        //  us bussiness area mein user nhi h
                        businessAreas[BA]?.user_info?.push({
                            userinfo: {
                                Title: item?.Title,
                                Id: item?.Id,
                                Email: item?.Email,
                                Role: item?.Role,
                            },
                            logs_list: item?.logs_list,
                            diliveralbe_list: [delivery],
                        });
                    }
                } else {

                    //agr bussines area exist hi nhi n=krta meain new bussiss area
                    businessAreas[BA] = {
                        bussiness_area: BA,
                        user_info: [{
                            userinfo: {
                                Title: item?.Title,
                                Id: item?.Id,
                                Email: item?.Email,
                                Role: item?.Role,
                            },
                            logs_list: item?.logs_list,
                            diliveralbe_list: [delivery],
                        }],
                    };
                }
            });
        });

        Object.keys(businessAreas).forEach((key) => {
            result.push(businessAreas[key]);
        });



        let all_bussiess_AreaList = []

        const calculateSum = (data) => {
            data.forEach((area) => {
                area.user_info.forEach((user) => {
                    user.diliveralbe_list.forEach((deliverable) => {
                        let phSum = 0;
                        let ahSum = 0;
                        deliverable.taskLists.forEach((task) => {
                            phSum += task.PlannedHours;
                            ahSum += task.ActualHours;
                        });

                        all_bussiess_AreaList.push({
                            BusinessArea: area.bussiness_area,
                            PH: phSum
                            ,
                            AH: ahSum
                        })
                    });
                });
            });
        };

        calculateSum(result);


        const businessAreas2 = {};

        all_bussiess_AreaList.forEach((obj) => {
            const { bussiness_area, PH, AH } = obj;

            if (businessAreas2.hasOwnProperty(bussiness_area)) {
                businessAreas2[bussiness_area].PH += PH;
                businessAreas2[bussiness_area].AH += AH;
            } else {
                businessAreas2[bussiness_area] = { PH, AH };
            }
        });
        result.sort((a, b) => a.bussiness_area.localeCompare(b.bussiness_area));
        setSlDisplayData(result)
        setIsSpinnerOn(false)
    }


    // fetch diliverable task 
    const getPbData = async (annualPlanItems) => {

        let annualPlanToBeReturnAlongWithTasks: any = []
        await sharepointWeb.lists
            .getByTitle("ProductionBoard")
            .items.select(
                "*,Developer/Title,Developer/Id,Developer/EMail,Product/Title,Product/ProductVersion,AnnualPlanID/Title,AnnualPlanID/ProjectVersion,FieldValuesAsText/StartDate,FieldValuesAsText/EndDate"
            )
            .expand("Developer,Product,AnnualPlanID,FieldValuesAsText")
            .top(5000)
            .get()
            .then(async (items) => {
                let _pbAllitems = [];
                let aplist = annualPlanItems;

                items.forEach((item, Index) => {
                    const eventStartDate = moment(item.StartDate, 'YYYY-MM-DD');

                    let aplIndex = aplist.findIndex(y => y.Id == item?.AnnualPlanIDId);
                    if (aplIndex !== -1
                        // &&
                        // eventStartDate.isSameOrAfter(firstDate, 'day') && eventStartDate.isSameOrBefore(lastDate, 'day')
                    ) {

                        let planned_hours = 0;
                        let allocated_hours = 0;
                        if (item.PlannedHours) {
                            planned_hours = item.PlannedHours;

                        }
                        if (item.ActualHours) {
                            allocated_hours = item.ActualHours;

                        }
                        aplist[aplIndex] = {
                            ...aplist[aplIndex],
                            diliverable_ph: aplist[aplIndex].diliverable_ph + planned_hours,
                            diliverable_ah: aplist[aplIndex].diliverable_ah + allocated_hours,
                            taskLists: [
                                ...aplist[aplIndex].taskLists,
                                {
                                    Tasks: item.Title,
                                    PlannedHours: planned_hours,
                                    ActualHours: allocated_hours,
                                }
                            ]
                        }
                    }


                });


                annualPlanToBeReturnAlongWithTasks = await getDpData(_pbAllitems, aplist);

            })
            .catch((error) => {
                console.log(error, "getPbData");
            });


        return annualPlanToBeReturnAlongWithTasks;
    };


    //fetch play book data
    const getDpData = async (data, annualPlanData) => {
        let finilisedReturnedItems = []
        await sharepointWeb.lists
            .getByTitle("Delivery Plan")
            .items.select(
                "*,Developer/Title,Developer/Id,Developer/EMail,Product/Title,Product/ProductVersion,AnnualPlanID/Title,AnnualPlanID/ProjectVersion,FieldValuesAsText/StartDate,FieldValuesAsText/EndDate"
            )
            .expand("Developer,Product,AnnualPlanID,FieldValuesAsText")
            // .filter("NotApplicable eq null or NotApplicable ne true")
            .top(5000)
            .get()
            .then(async (items) => {
                let _pbAllitems = data;
                let annualPlanDataLocal = annualPlanData;
                items.forEach((item) => {
                    const eventStartDate = moment(item.StartDate, 'YYYY-MM-DD');
                    let aplIndex = annualPlanDataLocal.findIndex(y => y.Id == item?.AnnualPlanIDId);
                    if (aplIndex !== -1 && item?.NotApplicable == null
                        // && eventStartDate.isSameOrAfter(firstDate, 'day') && eventStartDate.isSameOrBefore(lastDate, 'day')
                    ) {

                        annualPlanDataLocal[aplIndex] = {
                            ...annualPlanDataLocal[aplIndex], playBooks: [
                                ...annualPlanDataLocal[aplIndex].playBooks,
                                {
                                    Activity: item?.Title,
                                    StartDate: item?.StartDate,
                                    EndDate: item?.EndDate,
                                    Status: item?.Status,
                                }
                            ]
                        }
                    }


                });


                finilisedReturnedItems = [...annualPlanDataLocal];

            })
            .catch((error) => {
                console.log(error, "getDpData");
            });

        return finilisedReturnedItems;
    };

    const generateExcel = () => {
        const workbook = new Excel.Workbook();
        let listBA = businessWiseData;
        const worksheet2 = workbook.addWorksheet("BA Weekly Report");
        worksheet2.addRow([])
        worksheet2.columns = [
            { header: "", key: "listBA", width: 25 },
            { header: "", key: "phah", width: 25 },
        ];
        worksheet2.columns = [
            { header: "", key: "bA", width: 25 },

        ];
        //adding row header
        var row = worksheet2.addRow(["Business Area", "PH/AH"])

        //adding bg color to row header
        for (var index = 1; index <= 2; index++) {
            row.getCell(index).fill = {
                type: "pattern",
                pattern: "solid",
                fgColor: { argb: "00e8d1" },
            };
        }

        //addin valaue to bussines area


        listBA.forEach((item) => {
            worksheet2.addRow({
                listBA: item.businessArea ? item.businessArea : "",
                phah: (item?.plannedHours || item?.actualHours) ? item?.plannedHours.toFixed(2) + "/" + item?.actualHours.toFixed(2) : "",
            });
        });
        worksheet2.addRow([])
        let arrExport = DisplayData;


        arrExport.forEach((item) => {


            //empty row
            worksheet2.addRow([]);

            // bussiness diffrn row 
            worksheet2.addRow({

                bA: item.bussiness_area ? item.bussiness_area + ":" : ""
            })

            // var row2 = worksheet2.addRow(["First Name", "Deliverable Name", "Client", "Status", "Term", "Tasks", "PH/AH",
            //     //"Anomalies", "Requests", "Request Date", "Sent To", "Response Date", "Responses", "Activity", "Start Date", "End Date"
            // ])


            // // Apply the color to the specified cells
            // for (var index = 1; index <= 18; index++) {
            //     row2.getCell(index).fill = {
            //         type: "pattern",
            //         pattern: "solid",
            //         fgColor: { argb: "00e8d1" },
            //     };
            // }



            for (let x = 0; x < item?.user_info?.length; x++) {
                const user_info = item.user_info[x] || {}


                worksheet2.columns = [
                    { header: "", key: "firstName", width: 25 },
                    { header: "", key: "role", width: 25 },
                    { header: "", key: "delivName", width: 25 },
                    { header: "", key: "client", width: 25 },
                    { header: "", key: "status", width: 25 },
                    { header: "", key: "term", width: 25 },
                    { header: "", key: "tasks", width: 25 },
                    { header: "", key: "phah", width: 25 },
                    // { header: "", key: "anomalies", width: 25 },
                    // { header: "", key: "req", width: 25 },
                    // { header: "", key: "reqDate", width: 25 },
                    // { header: "", key: "sent", width: 25 },
                    // { header: "", key: "resDate", width: 25 },
                    // { header: "", key: "res", width: 25 },
                    // { header: "", key: "activity", width: 25 },
                    // { header: "", key: "startDate", width: 25 },
                    // { header: "", key: "endDate", width: 25 },
                ];

                var row2 = worksheet2.addRow(["First Name", "Deliverable Name", "Client", "Status", "Term", "Tasks", "PH/AH",
                    //"Anomalies", "Requests", "Request Date", "Sent To", "Response Date", "Responses", "Activity", "Start Date", "End Date"
                ])

                for (var index = 1; index <= 7; index++) {
                    row2.getCell(index).fill = {
                        type: "pattern",
                        pattern: "solid",
                        fgColor: { argb: "00e8d1" },
                    };
                }
                //all diliverable of a user
                for (let i = 0; i < user_info.diliveralbe_list.length; i++) {

                    const deliv = user_info.diliveralbe_list[i] || {};
                    /////////
                    const maxRows3 = Math.max(
                        // deliv?.playbooks?.length,
                        deliv?.taskLists?.length,
                        // deliv?.documents_review_log?.length, 
                        0

                    );

                    for (let z = 0; z <= maxRows3; z++) {

                        // const playbook = deliv.playbooks[z] || {};
                        const task = deliv.taskLists[z] || {};
                        // const log = deliv.documents_review_log[z] || {};
                        worksheet2.addRow({
                            firstName: user_info.userinfo.Title && z == 0 ? user_info.userinfo.Title : "",
                            // role: user_info.userinfo.Role && z == 0 ? user_info.userinfo.Role : "",
                            delivName: deliv.deliverable_name && z == 0 ? deliv.deliverable_name : "",
                            client: deliv.client && z == 0 ? deliv.client : "",
                            status: deliv.status && z == 0 ? deliv.status : "",
                            term: deliv.terms && z == 0 ? deliv.terms.join(", ") : "",
                            tasks: task.Tasks ? task.Tasks : " ",
                            phah: (task?.ActualHours || task?.PlannedHours) ? task?.ActualHours + "/" + task?.PlannedHours : " ",


                            // anomalies: log.anomilies ? log.anomilies : "",
                            // req: log.request ? log.request.join(", ") : "",
                            // reqDate: log.request_date ? moment(log.request_date).format(DateListFormat) : "",
                            // sent: log.sent_to ? log.sent_to.join(", ") : "",
                            // resDate: log.response_date ? moment(log.response_date).format(
                            //     DateListFormat
                            // ) : "",
                            // res: log.response ? log.response.join(", ") : "",
                            // activity: playbook.Activity ? playbook.Activity : "",
                            // startDate: playbook.StartDate ? moment(playbook.StartDate).format(
                            //     DateListFormat
                            // ) : "",
                            // endDate: playbook.EndDate ? moment(playbook.EndDate).format(
                            //     DateListFormat
                            // ) : "",
                        });
                    }


                }



                worksheet2.columns = [
                    // { header: "", key: "firstName", width: 25 },
                    // { header: "", key: "role", width: 25 },
                    // { header: "", key: "delivName", width: 25 },
                    // { header: "", key: "client", width: 25 },
                    // { header: "", key: "status", width: 25 },
                    // { header: "", key: "term", width: 25 },
                    // { header: "", key: "tasks", width: 25 },
                    // { header: "", key: "phah", width: 25 },
                    { header: "", key: "anomalies", width: 25 },
                    { header: "", key: "req", width: 25 },
                    { header: "", key: "reqDate", width: 25 },
                    { header: "", key: "sent", width: 25 },
                    { header: "", key: "resDate", width: 25 },
                    { header: "", key: "res", width: 25 },
                    // { header: "", key: "activity", width: 25 },
                    // { header: "", key: "startDate", width: 25 },
                    // { header: "", key: "endDate", width: 25 },
                ];

                //Header values
                var row2_2 = worksheet2.addRow([//"First Name", "Deliverable Name", "Client", "Status", "Term", "Tasks", "PH/AH",
                    "Anomalies", "Requests", "Request Date", "Sent To", "Response Date", "Responses"  //, "Activity", "Start Date", "End Date"
                ])

                for (var index = 1; index <= 7; index++) {
                    row2_2.getCell(index).fill = {
                        type: "pattern",
                        pattern: "solid",
                        fgColor: { argb: "00e8d1" },
                    };
                }




                let logs_data = [...reArrangeDataLog(item.user_info)]

                //all doc review logs of user
                const maxRows3_2 = Math.max(
                    // deliv?.playbooks?.length,
                    //    deliv?.taskLists?.length,
                    logs_data?.length,
                    0

                );

                for (let e = 0; e <= maxRows3_2; e++) {

                    const log = logs_data[e] || {};
                    worksheet2.addRow({
                        firstName: user_info.userinfo.Title && e == 0 ? user_info.userinfo.Title : "",
                        role: user_info.userinfo.Role && e == 0 ? user_info.userinfo.Role : "",



                        anomalies: log.anomilies ? log.anomilies : "",
                        req: log.request ? log.request.join(", ") : "",
                        reqDate: log.request_date ? moment(log.request_date).format(DateListFormat) : "",
                        sent: log.sent_to ? log.sent_to.join(", ") : "",
                        resDate: log.response_date ? moment(log.response_date).format(
                            DateListFormat
                        ) : "",
                        res: log.response ? log.response.join(", ") : "",

                    });
                }





                let play_data = [...reArrangeDataPlayBook(item.user_info)]

                //all doc review logs of user
                const maxRows3_3 = Math.max(
                    // deliv?.playbooks?.length,
                    //    deliv?.taskLists?.length,
                    play_data?.length,
                    0

                );

                for (let e = 0; e <= maxRows3_2; e++) {

                    const log = logs_data[e] || {};
                    worksheet2.addRow({
                        activity: log?.Activity ? log?.Activity : "",
                        startDate: log?.StartDate ? moment(log?.StartDate).format(
                            DateListFormat
                        ) : "",
                        endDate: log?.EndDate ? moment(log?.EndDate).format(
                            DateListFormat
                        ) : "",

                    });
                }
            }
            // Action Table
            var row3 = worksheet2.addRow(["Who", "Date", "Action Item",])


            // Apply the color to the specified cells
            for (var index = 1; index <= 3; index++) {
                row3.getCell(index).fill = {
                    type: "pattern",
                    pattern: "solid",
                    fgColor: { argb: "00e8d1" },
                };
            }
        });

        workbook.xlsx
            .writeBuffer()
            .then((buffer) =>
                FileSaver.saveAs(
                    new Blob([buffer]),
                    `BusinessArea-${new Date().toLocaleString()}.xlsx`
                )
            )
            .catch((err) => console.log("Error writing excel export", err));
    };
    const handleYearChange = (e: any) => {
        setSelectedYear(e);
        generateWeeksList(e)

    }
    const handleWeekChange = (e: any) => {
        if (Number(selectedYear) !== 0) {
            const weekNumber = parseInt(e.split(' ')[1]);
            // Calculate the first day of the selected week 
            const firstDayOfWeek = moment(selectedYear, 'YYYY')
                .startOf('isoWeek')
                .add(weekNumber, 'weeks');
            // Calculate the last day of the selected week
            // const lastDayOfWeek = firstDayOfWeek.clone().endOf('week'); 

            const lastDayOfWeek = moment(firstDayOfWeek).endOf('isoWeek');


            setFirstDate(moment(firstDayOfWeek, 'YYYY-MM-DD').format('YYYY-MM-DD'));
            setLastDate(moment(lastDayOfWeek, 'YYYY-MM-DD').format('YYYY-MM-DD'));
            setSelectedWeek(e);
        }

    }

    const isWithin5Percent = (ah, ph) => {
        const lowerBound = ah * 0.95;
        const upperBound = ah * 1.05;
        return ph >= lowerBound && ph <= upperBound;
    }



    const truncateString = (str, maxLength) => {
        if (str.length > maxLength) {
            return str.substring(0, maxLength) + '...';
        } else {
            return str;
        }
    }


    const reArrangeDataLog = (data) => {


        let logs = data?.flatMap(item =>
            item?.diliveralbe_list?.flatMap(deliverable =>
                deliverable?.documents_review_log
            )
        );

        // let playbooks = data.flatMap(item =>
        //     item.diliveralbe_list.flatMap(deliverable =>
        //         deliverable.playbooks
        //     )
        // );
        // console.log(logs);
        // const resp = {
        //     logs: logs,
        //     playbooks: playbooks
        // }
        return logs
    }
    const reArrangeDataPlayBook = (data) => {


        // let logs = data.flatMap(item =>
        //     item.diliveralbe_list.flatMap(deliverable =>
        //         deliverable.documents_review_log
        //     )
        // );

        let playbooks = data?.flatMap(item =>
            item?.diliveralbe_list?.flatMap(deliverable =>
                deliverable?.playbooks
            )
        );
        // const resp = {
        //     logs: logs,
        //     playbooks: playbooks
        // }
        return playbooks
    }


    useEffect(() => {
    }, [DisplayData])


    useEffect(() => {
        setDrLoader("startUpLoader");
        // queryGenerator(overallQueryArr);
        drGetAllOptions();
        YearFinder();
    }, [drReRender]);

    return <>
        <>
            <div style={{ padding: "5px 10px" }}>
                {drLoader == "startUpLoader" ? <CustomLoader /> : null}
                <div>
                    <div className={styles.dpTitle}
                        style={{
                            justifyContent: "space-between",
                            alignItems: "flex-start",
                            marginBottom: 10,
                        }}                    >
                        <Label style={{ color: "#000", fontSize: 24, padding: 0 }}>
                            
                        </Label>
                        <Label
                            onClick={() => {
                                generateExcel();
                            }}
                            style={{
                                backgroundColor: "#EBEBEB",
                                padding: "0 15px",
                                cursor: "pointer",
                                fontSize: "12px",
                                display: "flex",
                                alignItems: "center",
                                justifyContent: "center",
                                borderRadius: "3px",
                                color: "#1D6F42",
                                height: 34,
                            }}
                        >
                            <Icon
                                style={{
                                    color: "#1D6F42",
                                }}
                                iconName="ExcelDocument"
                            // className={apIconStyleClass.export}
                            />
                            Export as XLS
                        </Label>
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
                            flexWrap: "wrap",
                        }}
                    >
                    </div>
                </div>
                {/* Filter-Section Ends */}
                {/* Date Section */}
                <div style={{ display: "flex", flexWrap: "wrap", alignItems: "center", marginBottom: '10px' }}>

                    <div className="" >
                    <Label styles={drToggleLabelStyles}>{isToggleOn ? `Week Wise` : `Date Wise`}</Label>
                        <Stack tokens={stackTokens}>
                            <Toggle  defaultChecked onText="On" offText="Off" onChange={_onChangeToggle} />
                        </Stack></div>
                    {
                        isToggleOn ? <> <div className="" style={{ marginRight: '10px' }}>
                            <Label styles={drLabelStyles}>Year</Label>

                            <Dropdown
                                styles={
                                    selectedYear == "2023"
                                        ? drDropdownStyles
                                        : drActiveDropdownStyles
                                }
                                placeholder="Select a year"
                                
                                options={yearList.map((year) => ({ key: year, text: year }))}
                                onChange={(e, option) => handleYearChange(option.key)}
                                defaultSelectedKey={selectedYear}
                            /></div>
                            <div className="">
                            <Label styles={drLabelStyles}>Week</Label>
             
                                <Dropdown
                                    styles={
                                        selectedWeek == "Week 1"
                                            ? drDropdownStyles
                                            : drActiveDropdownStyles
                                    }
                                    placeholder="Select a week"
                                   
                                    options={weekList.map((week) => ({ key: week, text: week }))}
                                    onChange={(e, option) => handleWeekChange(option.key)}
                                    defaultSelectedKey={selectedWeek}
                                />
                            </div></> : <>
                            <div className="" style={{ marginLeft: "11px", marginTop: "5px" }}>
                                <div className={rootClass}>
                                    <DatePicker
                                        firstDayOfWeek={firstDayOfWeekForDatePicker}
                                        label="From Date"
                                        placeholder="Select a date..."
                                        ariaLabel="Select a date"
                                        // DatePicker uses English strings by default. For localized apps, you must override this prop.
                                        strings={defaultDatePickerStrings}
                                        onSelectDate={(e) => {
                                            const date = moment(e);
                                            const formattedDate = date.format('YYYY-MM-DD');
                                            setFirstDate(formattedDate)
                                        }}
                                    />
                                </div>
                            </div>
                            <div className="" style={{ marginLeft: "11px", marginTop: "5px" }}>
                                <div className={rootClass}>
                                    <DatePicker
                                        firstDayOfWeek={firstDayOfWeekForDatePicker}
                                        label="To Date"
                                        placeholder="Select a date..."
                                        ariaLabel="Select a date"
                                        // DatePicker uses English strings by default. For localized apps, you must override this prop.
                                        onSelectDate={(e) => {
                                            const date = moment(e);
                                            const formattedDate = date.format('YYYY-MM-DD');
                                            setLastDate(formattedDate)
                                        }}
                                        strings={defaultDatePickerStrings}
                                    />
                                </div>
                            </div></>
                    }


                    <div style={{ marginLeft: "8px" }}>
                        <Label
                            onClick={() => {
                                drGetAllOptions();
                            }}                        >
                            <Icon
                                iconName="Search"
                                title="Click to Generate Reports"
                                className={PLIconStyleClass.refresh}
                            />
                        </Label>
                    </div>
                    {isSpinnerOn && <div style={{ marginLeft: "15px", marginTop: "28px" }}>
                        <Stack tokens={tokens.sectionStack}>
                            <Stack {...rowProps} tokens={tokens.spinnerStack}>
                                <Spinner size={SpinnerSize.medium} />
                            </Stack>
                        </Stack>
                    </div>}
                </div>
                {
                    isAvailableData && <>
                        <div className="MyTableWPR" style={{ width: "30%" }}>
                            <table className="tableWPR tableWPR-bordered">
                                <thead>
                                    <tr>
                                        <th > Business Area</th>
                                        <th >PH/AH</th>
                                    </tr>
                                </thead>
                                <tbody>{
                                    businessWiseData.map((business, index) => {
                                        return (
                                            <tr key={index}>
                                                <td >{business.businessArea}</td>
                                                <td style={{ textAlign: "right" }} >{business.plannedHours.toFixed(2)}/{business.actualHours.toFixed(2)}</td>
                                            </tr>
                                        )
                                    })}
                                </tbody>
                            </table>
                        </div>
                        <div style={{ display: "flex" }}>
                            {/* DetailList-Section Starts */}
                            <div>
                                <div>
                                    <div>
                                        <div className=" ">
                                            {DisplayData.map(data => (
                                                <>
                                                    <h4 style={{ width: "300px !important" }}><u>{data.bussiness_area}</u>:</h4>
                                                    <div className="MyTableWPR" style={{ maxHeight: '800px', overflow: "auto" }}>
                                                        <table className="tableWPR tableWPR-bordered">
                                                            {
                                                                data.user_info.map((mou) => {
                                                                    return <>
                                                                        <thead>
                                                                            <tr>
                                                                                <th rowSpan={2} style={{ textAlign: 'left' }}> Name</th>
                                                                                <th colSpan={7}> Deliverable</th>
                                                                            </tr>
                                                                            <tr>
                                                                                <th>Deliverable Name</th>
                                                                                <th>Client</th>
                                                                                <th>Date</th>
                                                                                <th>Status</th>
                                                                                <th>Term</th>
                                                                                <th>Tasks</th>
                                                                                <th>PH/AH</th>
                                                                            </tr>
                                                                        </thead>
                                                                        <tbody>
                                                                            {/* {   [...reArrangeDataLog(data.user_info).logs].map((ex)=>{ */}
                                                                            {
                                                                                mou?.diliveralbe_list?.map((c: any, index) => {
                                                                                    return <>
                                                                                        <tr>
                                                                                            <td>{index === 0 ? mou?.userinfo.Title : <>&nbsp;</>}</td>
                                                                                            <td className={(c.diliverable_ph == 0 || c.diliverable_ah == 0 || isWithin5Percent(c.diliverable_ah, c.diliverable_ph) == true) && `annomolies-bg`}
                                                                                            >{c.deliverable_name}</td>
                                                                                            <td>{c.client}</td>
                                                                                            <td>{moment(c.date).format("DD-MM-YYYY")
                                                                                            }</td>
                                                                                            <td>{c.status}</td>
                                                                                            <td>{c?.terms?.join(", ")}</td>
                                                                                            <td>
                                                                                                {c?.taskLists.map((taskList) => (
                                                                                                    <tr style={{ border: "none", padding: 0 }}>
                                                                                                        {/* <TooltipHost content={taskList?.Tasks} delay={0}> */}
                                                                                                        <td style={{ borderLeft: " none", borderRight: "none", borderTop: "none", padding: 0 }}>
                                                                                                            {taskList?.Tasks}
                                                                                                        </td>
                                                                                                        {/* </TooltipHost> */}
                                                                                                    </tr>
                                                                                                ))}
                                                                                            </td>
                                                                                            <td>{c?.taskLists.map((taskList) => (
                                                                                                <tr style={{ border: "none", padding: 0 }}>
                                                                                                    <td style={{ borderLeft: " none", borderRight: "none", borderTop: "none", textAlign: "right", padding: 0 }}>
                                                                                                        {(taskList?.ActualHours || taskList?.PlannedHours) && (taskList?.PlannedHours + "/" + taskList?.ActualHours
                                                                                                            // .substring(0, 8) + '...'
                                                                                                        )}
                                                                                                    </td>
                                                                                                </tr>
                                                                                            ))}</td>
                                                                                            {/* <td>Role</td> */}
                                                                                        </tr>
                                                                                    </>
                                                                                })
                                                                            }

                                                                            {/* })} */}

                                                                        </tbody>
                                                                        <thead>
                                                                            {
                                                                                [...reArrangeDataLog(data?.user_info)]?.length != 0 && <>
                                                                                    <tr>
                                                                                        <th rowSpan={2} style={{ textAlign: 'left', backgroundColor: 'white', borderTop: 'none' }}> &nbsp; </th>
                                                                                        {/* <th rowspan="2">   Role</th> */}
                                                                                        {/* <th colspan="6">   Deliverable</th> */}
                                                                                        <th colSpan={7}>Document Submitted</th>
                                                                                        {/* <th colspan="4">PlayBook</th> */}
                                                                                    </tr>
                                                                                    <tr>
                                                                                        <th>Anomalies</th>
                                                                                        <th>Requests</th>
                                                                                        <th>Requests Date</th>
                                                                                        <th>Sent To</th>
                                                                                        <th>Response Date</th>
                                                                                        <th colSpan={2}>Response</th>
                                                                                    </tr></>
                                                                            }
                                                                        </thead>
                                                                        <tbody>
                                                                            {[...reArrangeDataLog(data?.user_info)]?.map((ex) => {
                                                                                return <tr>
                                                                                    <td> &nbsp; </td>
                                                                                    <td>{ex.anomilies}</td>
                                                                                    <td>{ex.request}</td>
                                                                                    <td>{
                                                                                        moment(ex.request_date).format("DD-MM-YYYY")
                                                                                    }</td>
                                                                                    <td>{ex.sent_to}</td>
                                                                                    <td>{moment(ex.response_date).format("DD-MM-YYYY")
                                                                                    }</td>
                                                                                    <td colSpan={2}>{ex.response}</td>
                                                                                    {/* <td>Role</td> */}
                                                                                </tr>
                                                                            })}


                                                                        </tbody>
                                                                        <thead>
                                                                            {
                                                                                [...reArrangeDataPlayBook(data?.user_info)]?.length != 0 && <>
                                                                                    <tr>
                                                                                        <th rowSpan={2} style={{ textAlign: 'left', backgroundColor: 'white', borderTop: 'none' }}> &nbsp; </th>
                                                                                        <th colSpan={7}>PlayBook</th>
                                                                                    </tr>
                                                                                    <tr>
                                                                                        <th>Activity</th>
                                                                                        <th>Start Date</th>
                                                                                        <th>End Date</th>
                                                                                        <th colSpan={4}>Status</th>
                                                                                    </tr>
                                                                                </>
                                                                            }
                                                                        </thead>
                                                                        <tbody>
                                                                            {[...reArrangeDataPlayBook(data?.user_info)]?.map((ex) => {
                                                                                return <tr>
                                                                                    <td> &nbsp; </td>
                                                                                    <td>{ex.Activity}</td>
                                                                                    <td>{moment(ex.StartDate).format("DD-MM-YYYY")

                                                                                    }</td>
                                                                                    <td>{
                                                                                        moment(ex.EndDate).format("DD-MM-YYYY")
                                                                                    }</td>
                                                                                    <td colSpan={4}>{ex.Status}</td>
                                                                                </tr>
                                                                            })}

                                                                        </tbody>
                                                                    </>
                                                                })
                                                            }

                                                        </table>
                                                    </div>
                                                </>))}
                                        </div>
                                    </div >
                                </div>
                            </div>

                            {/* Popup-Section Ends */}
                        </div>
                    </>
                }

            </div>
        </>
    </>

};

export default BusinessArea;
