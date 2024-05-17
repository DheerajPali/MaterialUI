import * as React from 'react';
import type { IMaterialUiGrid1Props } from './IMaterialUiGrid1Props';
import { ImaterialGridState } from './IMaterialGrid1State';
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/items/get-all";
import "@pnp/sp/fields";
import "@pnp/sp/attachments";
import "@pnp/sp/webs";
import "@pnp/sp/folders/web";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import { SPFx, spfi } from '@pnp/sp';

export default class MaterialUiGrid1 extends React.Component<IMaterialUiGrid1Props, ImaterialGridState> {
  constructor(props: IMaterialUiGrid1Props) {
    super(props)
    this.state = {
      InvoiceNo: NaN,
      CompanyName: '',
      Invoicedetails: '',
      CompanyCode: '',
      InvoiceAmount: NaN,
      BasicValue: NaN,
      Approver: [],
      IsApproved: false,
      Country: '',
      IsDeleted: false,
      data: [],
    }
  }
  //This method runs when component mount(or come into our DOM)
  public componentDidMount = async () => {
    await this.getAll();
  }
  //This method fetch all the records available inside list. 
  public getAll = async () => {
    try {
      const val = false;
      const sp: any = spfi().using(SPFx(this.props.context));
      const listData = await sp.web.lists.getByTitle("InvoiceDetails").items.filter(`IsDeleted eq ${val}`).select(
        "InvoiceNo",
        "CompanyName",
        "Invoicedetails",
        "CompanyCode",
        "InvoiceAmount",
        "BasicValue",
        "Approver/Title",
        "IsApproved",
        "Country",
      ).expand("Approver")();
      this.setState({
        data: listData,
      })
    } catch (error) {
      console.log("error : ", error);
    }
  }

  public render(): React.ReactElement<IMaterialUiGrid1Props> {
    return (
      <>
        {this.state.data.map((item: any) => {
          return (
            <div>
              <div>{item.InvoiceNo}</div>
              <div>{item.CompanyName}</div>
              <div>{item.Invoicedetails}</div>
              <div>{item.CompanyCode}</div>
              <div>{item.InvoiceAmount}</div>
              <div>{item.BasicValue}</div>
              <div>{item.IsApproved}</div>
              <div>{item.Country}</div>
            </div>
          )
        })}
      </>
    );
  }
}


