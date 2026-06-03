import * as React from "react";
import "./userDashboardsc.scss";
import APperformerAdvanceFormForUTR from "./APperformerAdvanceFormForUTR";

import { useState } from "react";
import APperformerAdvanceform from "./APperformerAdvanceform";
import logo from "../assets/SonaPNGLogo.png";
import Edit from "../assets/Pencil.png";
import User from "../assets/Userlogo.png";
//import "../assets/bootstrap/css/bootstrap.css";

import { spfi } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";

interface UserDashboardProps {
  context: any;
}

const APperformerDashboard: React.FC<UserDashboardProps> = ({ context }) => {
  const sp = spfi().using(SPFx(context));
  //const [formType, setFormType] = useState<"new" | "view" | null>(null);
  //const [formType, setFormType] = useState<"new" | "view" | "approve" | null>(null);

  const [formType, setFormType] = useState<
    "approve" | "approveUTR" | "view" | null
  >(null);

  const [activeMenu, setActiveMenu] = React.useState("My Request");
  const [searchText, setSearchText] = React.useState("");
  const [statusFilter, setStatusFilter] = React.useState("");
  const [showForm, setShowForm] = React.useState(false);
  const [data, setData] = React.useState<any[]>([]);
  const [currentUserName, setCurrentUserName] = React.useState("");
  const [selectedItem, setSelectedItem] = React.useState<any>(null);

  // const username = props.userDisplayName;
  const handleFormOpen = async (
    item: any,
    type: "view" | "approve" | "approveUTR",
  ) => {
    try {
      debugger;
      const fullItem = await sp.web.lists
        .getByTitle("CapexPayment")
        .items.getById(item.ID)
        .select(
          "*",
          "ID",
          "Title",
          "EmployeeCode",
          "EmployeeName",
          "Email",
          "Division",
          "Location",
          "ContactNo",
          "EmployeeStatus",

          "RM",
          "HOD",

          "CapexId",
          "VendorName",
          "VendorCode",

          "PONumber",
          "PODate",
          "POAmount",
          "POPaymentTerms",

          "MRNNumber",
          "MRNDtae",
          "MRNAmountwithGST",

          "RequestedAmountforPayment",
          "VoucherNumber",
          "VoucherDate",
          "ApproverRemarks",
          "FinalPaymentAgainstPO",
          "InstallationDetails",

          "Status",
          "Author/EMail",
        )
        .expand("Author")();

      // ✅ Map SharePoint fields to UI fields
      const mappedData = {
        ...fullItem,
        EmployeeEmail: fullItem.EmployeeEmail,
        POAmtGST: fullItem.POAmount,
        POAdvanceTerms: fullItem.POPaymentTerms,

        mrnNumber: fullItem.MRNNumber,
        mrnDate: fullItem.MRNDtae,
        mrnAmount: fullItem.MRNAmountwithGST,

        requestedAmount: fullItem.RequestedAmountforPayment,

        finalPayment: fullItem.FinalPaymentAgainstPO,
        installationDetails: fullItem.InstallationDetails,
        ReportingManager: fullItem.ReportingManager,
        HOD: fullItem.HOD,
        ApprovalMatrix: fullItem.ApprovalMatrix ? JSON.parse(fullItem.ApprovalMatrix) : null,
        WorkflowHistory: fullItem.WorkflowHistory ? JSON.parse(fullItem.WorkflowHistory) : null,
      };

      setSelectedItem(mappedData);
      setFormType(type);
      setShowForm(true);
    } catch (error) {
      console.error("Form open error:", error);
    }
  };
  // ✅ GET CURRENT USER
  const getLoggedInUser = async () => {
    try {
      const user = await sp.web.currentUser();
      setCurrentUserName(user.Title);
    } catch (error) {
      console.error("User error:", error);
    }
  };
  const handleApproveClick = async (item: any) => {
    try {
      const fullItem = await sp.web.lists
        .getByTitle("CapexPayment")
        .items.getById(item.ID)
        .select("*", "PICName/Title")
        .expand("PICName")();

      setSelectedItem(fullItem);

      // 🔥 CONDITION BASED ROUTING
      if (item.status === "Pending for PF Approver") {
        setFormType("approve");
      } else if (item.status === "Pending for PF Approver UTR") {
        setFormType("approveUTR");
      }

      setShowForm(true);
    } catch (error) {
      console.error("Approve error:", error);
    }
  };
  const getCapexData = async () => {
    try {
      const currentUser = await sp.web.currentUser();
      let filterQuery;
      // const currentUser = await sp.web.currentUser();

      if (activeMenu === "My Request") {
        filterQuery = `CurrentApproverId eq ${currentUser.Id} and Status eq 'Pending for PF Approver' or Status eq 'Pending for PF Approver UTR'`;
      } else if (activeMenu === "Paid") {
        filterQuery = `CurrentApproverId eq ${currentUser.Id} and Status eq 'Paid'`;
      } else if (activeMenu === "Rejected") {
        filterQuery = `CurrentApproverId eq ${currentUser.Id} and Status eq 'Rejected'`;
      }
      const items = await sp.web.lists
        .getByTitle("CapexPayment")
        .items.select(
          "*",
          "ID",
          "Title",
          "Created",
          "EmployeeName",
          "VendorName",
          "VendorCode",
          "PONumber",
          "RequestedAmountforPayment",
          "Status",
          "Author/EMail", // 🔥 needed for My Request
        )
        .expand("Author") // ✅ correct expand
        .filter(filterQuery)
        .orderBy("ID", false)();

      const formatted = items.map((item: any) => ({
        ID: item.ID,
        id: item.Title,
        date: item.Created
          ? new Date(item.Created).toLocaleDateString("en-GB")
          : "",
        EmployeeName: item.EmployeeName || "",
        vendor: item.VendorName || "",
        vendorCode: item.VendorCode || "",
        po: item.PONumber || "",
        amount: item.RequestedAmountforPayment || 0,
        status: item.Status || "", // ✅ FIXED
      }));

      setData(formatted);
    } catch (error) {
      console.error("Data error:", error);
    }
  };

  const filteredData = data.filter((item) => {
    const text = searchText.toLowerCase();
    const status = statusFilter.toLowerCase();

    let menuFilter = true;

    if (activeMenu === "Paid") {
      menuFilter = item.status?.toLowerCase() === "paid";
    } else if (activeMenu === "Rejected") {
      menuFilter = item.status?.toLowerCase() === "rejected";
    } else if (activeMenu === "My Request") {
      menuFilter = true;
    }

    return (
      menuFilter &&
      (item.id?.toLowerCase().includes(text) ||
        item.vendor?.toLowerCase().includes(text) ||
        item.po?.toLowerCase().includes(text)) &&
      (!status || item.status?.toLowerCase().includes(status))
    );
  });
  React.useEffect(() => {
    if (!context) return;
    void getLoggedInUser();
    void getCapexData();
  }, [context]);

  if (showForm && selectedItem) {
    if (formType === "approve") {
      return (
        <APperformerAdvanceform
          context={context}
          formData={selectedItem}
          onClose={() => setShowForm(false)}
        />
      );
    }

    if (formType === "approveUTR") {
      return (
        <APperformerAdvanceFormForUTR
          context={context}
          formData={selectedItem}
          onClose={() => setShowForm(false)}
        />
      );
    }
  }

  return (
    <div style={{ display: "flex", width: "100%" }}>
      <div className="sidebar">
        <div className="sidehead">
          <div className="logo">
            <img src={logo} width="25px" height="25px" />
          </div>
          <div className="sidehead-right">SONA COMSTAR</div>
        </div>

        <div className="sidehead-user">
          <img
            src={User}
            style={{ margin: "10px 20px" }}
            width={20}
            height={20}
          />
          {currentUserName}
        </div>

        <ul className="nav">
          <li className="nav-item">
            <a
              className={
                activeMenu === "My Request" ? " nav-link active" : "nav-link"
              }
              onClick={() => setActiveMenu("My Request")}
              style={{ cursor: "pointer" }}
            >
              My Request
            </a>
          </li>
          <li className="nav-item">
            <a
              className={
                activeMenu === "Paid" ? " nav-link  active" : "nav-link"
              }
              onClick={() => setActiveMenu("Paid")}
              style={{ cursor: "pointer" }}
            >
              Paid
            </a>
          </li>
          <li className="nav-item">
            <a
              className={
                activeMenu === "Rejected" ? "nav-link  active" : "nav-link"
              }
              onClick={() => setActiveMenu("Rejected")}
              style={{ cursor: "pointer" }}
            >
              Rejected
            </a>
          </li>
        </ul>
      </div>
      <div
        className="main"
        style={{ width: "calc(100% - 250px)", transition: "width 0.3s" }}
      >
        <div className="header">
          <div className="left-banner">
            <div className="logo-text">
              <h2> Capex Payment Performer Dashboard </h2>
            </div>
          </div>
        </div>
        <div className="col-md-12 mainsecond">
          <div>
            <input
              placeholder="Search"
              value={searchText}
              className="form-control"
              style={{ width: "250px;" }}
              onChange={(e) => setSearchText(e.target.value)}
            />
          </div>
          <div>
            <select
              value={statusFilter}
              className="formtext-control"
              onChange={(e) => setStatusFilter(e.target.value)}
            >
              <option value="">All</option>
              <option value="Submitted">Submitted</option>
              <option value="Approved">Approved</option>
              <option value="Rejected">Rejected</option>
              <option value="Draft">Draft</option>
            </select>
          </div>
        </div>
        <main className="Main-Dash mx-2">
          <div style={{ overflowX: "auto" }}>
            <div className="table-vert-scroll">
              <table className="custom-table min-w-full bg-white rounded-2xl shadow-md">
                <thead
                  className="text-white"
                  style={{ backgroundColor: "rgb(60, 62, 69)" }}
                >
                  <tr>
                    <th className="px-4 py-2">Action</th>
                    <th className="px-4 py-2">Payment ID</th>
                    <th className="px-4 py-2">Requestor Date</th>
                    <th className="px-4 py-2">Requestor Name</th>
                    <th className="px-4 py-2">Requestor Type</th>
                    <th className="px-4 py-2">Vendor Code</th>
                    <th className="px-4 py-2">Vendor Name</th>
                    <th className="px-4 py-2">PO Number</th>
                    <th className="px-4 py-2">Request Amount</th>

                    <th className="px-4 py-2">Pending With</th>
                    <th className="px-4 py-2">Status</th>
                    
                  </tr>
                </thead>
                <tbody>
                  {filteredData.length === 0 ? (
                    <tr>
                      <td colSpan={11} style={{ textAlign: "center" }}>
                        No Data
                      </td>
                    </tr>
                  ) : (
                    filteredData.map((item, i) => (
                      <tr key={i}>
                        <td className="px-4 py-2">
                          {(item.status === "Pending for PF Approver" ||
                            item.status === "Pending for PF Approver UTR") && (
                              <span
                                onClick={() => {
                                  if (item.status === "Pending for PF Approver") {
                                    handleFormOpen(item, "approve");
                                  } else if (
                                    item.status === "Pending for PF Approver UTR"
                                  ) {
                                    handleFormOpen(item, "approveUTR");
                                  }
                                }}
                                style={{ cursor: "pointer" }}
                              >
                                <img src={Edit} width={15} alt="Action" />
                              </span>
                            )}
                        </td>
                        <td className="px-4 py-2">{item.id}</td>
                        <td className="px-4 py-2">{item.date}</td>
                        <td className="px-4 py-2">{item.EmployeeName}</td>
                        <td className="px-4 py-2">Capex Payment</td>
                        <td className="px-4 py-2"> {item.vendorCode}</td>
                        <td className="px-4 py-2">{item.vendor}</td>
                        <td className="px-4 py-2">{item.po}</td>
                        <td className="px-4 py-2">₹ {item.amount}</td>

                        <td className="px-4 py-2">Approver</td>
                        <td className="px-4 py-2">{item.status}</td>
                        
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </main>
      </div>
    </div>
  );
};

export default APperformerDashboard;
