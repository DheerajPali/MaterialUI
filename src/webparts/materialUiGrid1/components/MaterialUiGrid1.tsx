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
import { MaterialReactTable } from "material-react-table";
import { Box, Button } from "@mui/material";
import EditIcon from "@mui/icons-material/Edit";
import DeleteIcon from "@mui/icons-material/Delete";
import { CSVLink } from "react-csv";

export default class MaterialUiGrid1 extends React.Component<IMaterialUiGrid1Props, ImaterialGridState> {
  public siteUrl: any = this.props.context.pageContext.web.absoluteUrl;
  public relativeUrl = "/SitePages/Invoice.aspx";
  public csvExporter: any;
  constructor(props: IMaterialUiGrid1Props) {

    super(props)
    const headerColumn: any = [
      {
        header: "Actions",
        accessorKey: "Actions",
        size: 110,
        muiTableBodyCellProps: {
          align: "center",
        },
        Cell: ({ row }: any) => (
          <Box>
            <div style={{ display: 'flex', justifyContent: 'space-evenly' }}>
              <div data-interception="off"  onClick={() => this.handleEdit(row.original.Id)}>
              <EditIcon />
              </div>
              {/* <div onClick={() => this.hanldeDeleteRecord(row.original.Id)}>
                <DeleteIcon />
              </div> */}
              <div onClick={() => this.handleSoftDelete(row.original.Id)}>
                <DeleteIcon />
              </div>
            </div>
          </Box>
        ),
        enableColumnFilter: true,
        enableSorting: true,
        enableGrouping: true,
      },
      {
        header: "Invoice No",
        accessorKey: "InvoiceNo",
        size: 120,
      },
      {
        header: "Company Name",
        accessorKey: "CompanyName",
        size: 120,
      },
      {
        header: "Invoice Details",
        accessorKey: "Invoicedetails",
        size: 120,
      },
      {
        header: "Company Code",
        accessorKey: "CompanyCode",
        size: 120,
      },
      {
        header: "Invoice Amount",
        accessorKey: "InvoiceAmount",
        size: 120,
      },
      {
        header: "Basic Value",
        accessorKey: "BasicValue",
        size: 120,
      },
      {
        header: "Country",
        accessorKey: "Country",
        size: 120,
      },
      // {
      //   header: 'IsApproved',
      //   accessorKey: 'IsApproved',
      //   size: 120,
      // },
      {
        header: "IsApproved",
        accessorKey: "IsApproved",
        size: 110,
        muiTableBodyCellProps: {
          align: "center",
        },
        Cell: ({ row }: any) => {
          return (
            <div
              style={{
                backgroundColor: row.original.IsApproved ? "Green" : "Red",
                border: row.original.IsApproved
                  ? "3px solid Green"
                  : "3px solid Red",
                color: "white",
                borderRadius: "10%",
                width: "fit-content",
              }}
            >
              {row.original.IsApproved ? "Yes" : "No"}
            </div>
          );
        },
      },
      {
        header: "Approver",
        accessorKey: "Approver",
        size: 110,
        muiTableBodyCellProps: {
          align: "center",
        },
        Cell: ({ row }: any) => (
          <div>
            {row.original.Approver ? (
              row.original.Approver.map((item: any) => {
                return (
                  <>
                    <span>{item.Title}</span>
                    <br />
                  </>
                );
              })
            ) : (
              <div>--</div>
            )}
          </div>
        ),
        enableColumnFilter: true,
        enableSorting: true,
        enableGrouping: true,
      },
    ];

    this.state = {
      ID: NaN,
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
      headerCols: headerColumn,
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
        "ID",
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
  public handleSoftDelete = async (id: number) => {
    try {
      const sp: any = spfi().using(SPFx(this.props.context));
      await sp.web.lists
        .getByTitle("InvoiceDetails")
        .items.getById(id)
        .update({
          IsDeleted: true,
        });
      alert('Deleted Successfully');
      this.getAll();
    } catch (error) {
      console.log("handleSoftDelete :: Error : ", error);
    }
  }

  public handleEdit = async (id: number) => {
    try {
      // Construct the URL with the item ID and mode type set to 'Edit'
      var editUrl;
      editUrl = `${this.siteUrl}${this.relativeUrl}?itemID=${id}`;
      // Open the edit URL in a new window
      window.open(editUrl, '_blank');
    } catch (error) {
      console.log("handleEdit :: Error :", error);
    }
  };  
  public render(): React.ReactElement<IMaterialUiGrid1Props> {
    const excelData = this.state.data.map((item: any) => ({
      ID: item.ID,
      InvoiceNo: item.InvoiceNo,
      CompanyName: item.CompanyName,
      Invoicedetails: item.Invoicedetails,
      CompanyCode: item.CompanyCode,
      InvoiceAmount: item.InvoiceAmount,
      BasicValue: item.BasicValue,
      Country: item.Country,
      IsApproved: item.IsApproved,
      Approver: item.Approver ? item.Approver.map((approver: { Title: string; }) => approver.Title).join(', ') : '',
    }));

    return (
      <>
        <MaterialReactTable
          displayColumnDefOptions={{
            "mrt-row-actions": {
              muiTableHeadCellProps: {
                align: "center",
              },
              size: 120,
            },
          }}
          columns={this.state.headerCols}
          data={this.state.data}
          // state={{ isLoading: true }}
          enableColumnResizing
          initialState={{
            density: "compact",
            pagination: { pageIndex: 0, pageSize: 100 },
            showColumnFilters: true,
          }}
          columnResizeMode="onEnd"
          positionToolbarAlertBanner="bottom"
          enablePinning
          // enableRowActions
          // onEditingRowSave={this.handleSaveRowEdits}
          // onEditingRowCancel={this.handleCancelRowEdits}
          enableGrouping
          enableStickyHeader
          enableStickyFooter
          enableDensityToggle={false}
          enableExpandAll={false}
          renderTopToolbarCustomActions={({ table }) => (
            <Box
              sx={{
                display: "flex",
                gap: "1rem",
                p: "0.5rem",
                flexWrap: "wrap",
              }}
            >

              <CSVLink
                data={excelData}
                filename={"InvoiceDetails.csv"} // Here you can provide a custom file name
                ref={(r: any) => (this.csvExporter = r)}
                style={{ display: "none" }} // Hide this link
              >
                Hidden Download me
              </CSVLink>
              <Button
                variant="contained"
                color="primary"
                onClick={() => {
                  this.csvExporter.link.click();
                }}
              >
                Export To Excel
              </Button>
            </Box>
          )}
        />
        {/* {this.state.data.map((item: any) => {
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
        })} */}
      </>
    );
  }
}


