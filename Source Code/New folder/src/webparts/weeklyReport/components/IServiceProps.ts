export interface IFilter
{
    FilterKey:string;
    FilterValue:string;
    Operator:string;
}
  
export interface IListItems
{
   Listname:string;
   Select?:string;
   Topcount?:number;
   Expand?:string;
   Orderby?:string;
   Orderbydecorasc?:boolean;
   Filter?:IFilter[];
   PageCount?:number;
   PageNumber?:number;
}

export interface IAddList
{
  Listname:string;
  RequestJSON:object;
}

export interface IDeleteList
{
  Listname:string;
  ID:number;
}

export interface IUpdateList
{
  Listname:string;
  RequestJSON:object;
  ID:number;
}