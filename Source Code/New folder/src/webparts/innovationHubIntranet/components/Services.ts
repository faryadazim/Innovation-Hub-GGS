import * as React from "react";
import { useState, useEffect } from "react";
import { sp } from "@pnp/sp/presets/all";
import * as moment from "moment";
import {
  IFilter,
  IListItems,
  IAddList,
  IUpdateList,
  IDeleteList,
  IDetailsListGroup,
  IDisplayDate,
} from "./IServiceProps";
import { IItemAddResult } from "@pnp/sp/items";
import "@pnp/sp/webs";
import { Web } from "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import "@pnp/sp/site-groups/web";

var webURL;

if (window.location.href.toLowerCase().indexOf("production/") > -1) {
  /* Production URL */
  webURL = "https://ggsaus.sharepoint.com";
} else {
  /* Development URL */
  webURL = "https://ggsaus.sharepoint.com/sites/Intranet_Test";
}

const _webURL = Web(webURL);

const SPAddItem = async (params: IAddList): Promise<IItemAddResult> => {
  return await _webURL.lists
    .getByTitle(params.Listname)
    .items.add(params.RequestJSON);
};
const SPUpdateItem = async (params: IUpdateList): Promise<IItemAddResult> => {
  return await _webURL.lists
    .getByTitle(params.Listname)
    .items.getById(params.ID)
    .update(params.RequestJSON);
};
const SPDeleteItem = async (params: IDeleteList): Promise<void> => {
  return await _webURL.lists
    .getByTitle(params.Listname)
    .items.getById(params.ID)
    .delete();
};
const SPReadItems = async (params: IListItems): Promise<[]> => {
  params = formatInputs(params);
  let filterValue: string = formatFilterValue(params.Filter);

  return await _webURL.lists
    .getByTitle(params.Listname)
    .items.select(params.Select)
    .filter(filterValue)
    .expand(params.Expand)
    .top(params.Topcount)
    .orderBy(params.Orderby, params.Orderbydecorasc)
    .get();
};

const SPDetailsListGroupItems = async (params: IDetailsListGroup) => {
  let newRecords = [];
  params.Data.forEach((arr, index) => {
    newRecords.push({
      Lesson: arr[params.Column],
      indexValue: index,
    });
  });

  let varGroup = [];
  let UniqueRecords = newRecords.reduce(function (item, e1) {
    var matches = item.filter(function (e2) {
      return e1[params.Column] === e2[params.Column];
    });

    if (matches.length == 0) {
      item.push(e1);
    }
    return item;
  }, []);

  UniqueRecords.forEach((ur) => {
    let recordLength = newRecords.filter((arr) => {
      return arr[params.Column] == ur[params.Column];
    }).length;
    varGroup.push({
      key: ur[params.Column],
      name: ur[params.Column],
      startIndex: ur.indexValue,
      count: recordLength,
    });
  });
  // console.log([...varGroup]);
  return varGroup;
};
const SPDisplayDate = (params: IDisplayDate) => {
  let TimeZone = params.TimeZone;
  var convertedDate = params.Date;
  var newTime = "";
  if (convertedDate) {
    if (TimeZone.includes("+")) {
      var Hourandmin = TimeZone.split("+")[1].split(":");
      newTime = moment(convertedDate, "YYYY-MM-DDTHH:mm")
        .add("hours", Hourandmin[0])
        .add("minutes", Hourandmin[1])
        .format("YYYY-MM-DD HH:mm");
    } else if (TimeZone.includes("-")) {
      var Hourandmin = TimeZone.split("-")[1].split(":");
      newTime = moment(convertedDate, "YYYY-MM-DDTHH:mm")
        .subtract("hours", Hourandmin[0])
        .subtract("minutes", Hourandmin[1])
        .format("YYYY-MM-DD HH:mm");
    }
  }
  return newTime;
  /*
    var convertedDate = moment(params.Date).format("DD/MM/YYYY");
    var newTime = "";
    if (convertedDate) {
      if (convertedDate.includes("+")) {
        var Hourandmin = convertedDate.split("+")[1].split(":");
        newTime = moment(new Date(convertedDate).toISOString())
          .add("hours", Hourandmin[0])
          .add("minutes", Hourandmin[1])
          .format("YYYY-MM-DD");
      } else if (params.Date.includes("-")) {
        var Hourandmin = convertedDate.split("-")[1].split(":");
        newTime = moment(new Date(convertedDate).toISOString())
          .subtract("hours", Hourandmin[0])
          .subtract("minutes", Hourandmin[1])
          .format("YYYY-MM-DD");
      }
    }
    return newTime;
    */
};

const readItemsFromSharepointListForDashbaord = async (
  params: IListItems
): Promise<[]> => {
  params = formatInputs(params);
  let filterValue: string = formatFilterValue(params.Filter);
  let skipcount = params.PageNumber * params.PageCount - params.PageCount;

  return await _webURL.lists
    .getByTitle(params.Listname)
    .items.select(params.Select)
    .filter(filterValue)
    .expand(params.Expand)
    .skip(skipcount)
    .top(params.PageCount)
    .orderBy(params.Orderby, params.Orderbydecorasc)
    .get();
};

const batchInsert = async (params: {
  ListName: string;
  responseData: any[];
}): Promise<any> => {
  const list = _webURL.lists.getByTitle(params.ListName);
  const batch = _webURL.createBatch();
  const promises: any[] = [];
  for (const data of params.responseData) {
    const promise = list.items.inBatch(batch).add(data);
    promises.push(promise);
  }
  await batch
    .execute()
    .then(() => {
      return promises;
    })
    .catch((error) => console.log(error));
};
const batchUpdate = async (params: {
  ListName: string;
  responseData: any[];
}): Promise<any> => {
  const list = _webURL.lists.getByTitle(params.ListName);
  const batch = _webURL.createBatch();
  const promises = [];
  for (const data of params.responseData) {
    const promise = list.items.getById(data.ID).inBatch(batch).update(data);
    promises.push(promise);
  }
  await batch
    .execute()
    .then(() => {
      return promises;
    })
    .catch((error) => console.log(error));
};

const formatInputs = (data: IListItems): IListItems => {
  !data.Select ? (data.Select = "*") : "";
  !data.Topcount ? (data.Topcount = 5000) : "";
  !data.Orderby ? (data.Orderby = "ID") : "";
  !data.Expand ? (data.Expand = "") : "";
  !data.Orderbydecorasc ? (data.Orderbydecorasc = true) : "";
  !data.PageCount ? (data.PageCount = 10) : "";
  !data.PageNumber ? (data.PageNumber = 1) : "";

  return data;
};
const formatFilterValue = (params: IFilter[]): string => {
  let strFilter: string = "";
  for (let i = 0; i < params.length; i++) {
    if (params[i].FilterKey) {
      if (i != 0) {
        strFilter += " and ";
      }

      if (
        params[i].Operator.toLocaleLowerCase() == "eq" ||
        params[i].Operator.toLocaleLowerCase() == "neq" ||
        params[i].Operator.toLocaleLowerCase() == "gt" ||
        params[i].Operator.toLocaleLowerCase() == "lt" ||
        params[i].Operator.toLocaleLowerCase() == "ge" ||
        params[i].Operator.toLocaleLowerCase() == "le"
      )
        strFilter +=
          params[i].FilterKey +
          " " +
          params[i].Operator +
          "'" +
          params[i].FilterValue +
          "'";
      else if (params[i].Operator.toLocaleLowerCase() == "substringof")
        strFilter +=
          params[i].Operator +
          "('" +
          params[i].FilterKey +
          "','" +
          params[i].FilterValue +
          "')";
    }
  }
  return strFilter;
};

export default {
  SPAddItem,
  SPUpdateItem,
  SPDeleteItem,
  SPReadItems,
  SPDetailsListGroupItems,
  SPDisplayDate,
  batchInsert,
  batchUpdate,
};
