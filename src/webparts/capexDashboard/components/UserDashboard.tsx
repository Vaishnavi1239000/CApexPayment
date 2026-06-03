import * as React from "react";
import "./userDashboardsc.scss";
import NewAdvanceform from "./NewAdvanceform";
import ViewAdvanceForm from "./ViewAdvanceForm";
import EditAdvanceForm from "./EditAdvanceForm";

import { useState } from "react";

import logo from "../assets/SonaPNGLogo.png";
import Edit from "../assets/Pencil.png";
import View from "../assets/Eye.png";
import User from "../assets/Userlogo.png";

import { spfi } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";

interface UserDashboardProps {
  context: any;
}

const UserDashboard: React.FC<UserDashboardProps> = ({ context }) => {
  const sp = spfi().using(SPFx(context));
  const [formType, setFormType] = useState<"view" | "new" | "edit" | null>(
    null,
  );

  const [activeMenu, setActiveMenu] = React.useState("My Request");
  const [searchText, setSearchText] = React.useState("");
  const [statusFilter, setStatusFilter] = React.useState("");
  const [showForm, setShowForm] = React.useState(false);
  const [data, setData] = React.useState<any[]>([]);
  const [currentUserName, setCurrentUserName] = React.useState("");
  const [selectedItem, setSelectedItem] = React.useState<any>(null);

  // ✅ GET CURRENT USER
  const getLoggedInUser = async () => {
    try {
      const user = await sp.web.currentUser();
      setCurrentUserName(user.Title);
    } catch (error) {
      console.error("User error:", error);
    }
  };

  // ✅ GET LIST DATA
  const getCapexData = async () => {
  try {
    debugger;

    let filterQuery = "";

    const currentUser = await sp.web.currentUser();

    // ✅ My Request
    if (activeMenu === "My Request") {
      filterQuery = `AuthorId eq ${currentUser.Id} and Status ne 'Paid' and Status ne 'Rejected'`;
    }

    // ✅ Paid
    else if (activeMenu === "Paid") {
      filterQuery = "Status eq 'Paid'";
    }

    // ✅ Rejected
    else if (activeMenu === "Rejected") {
      filterQuery = "Status eq 'Rejected'";
    }

    const items = await sp.web.lists
      .getByTitle("CapexPayment")
      .items.select(
        "ID",
        "Title",
        "Created",
        "EmployeeName",

        // ✅ SIMPLE TEXT FIELDS
       // "VendorCode",
        "VendorName",

        "PONumber",
        "RequestedAmountforPayment",
        "Status",

       
      )
      
      .filter(filterQuery)
      .orderBy("ID", false)();

    console.log("Items:", items);

    const formatted = items.map((item: any) => ({
      ID: item.ID,

      id: item.Title,

      date: item.Created
        ? new Date(item.Created).toLocaleDateString("en-GB")
        : "",

      EmployeeName: item.EmployeeName || "",

      // ✅ DIRECT FIELDS
      vendor: item.VendorName || "",

     // vendorCode: item.VendorCode || "",

      po: item.PONumber || "",

      amount: item.RequestedAmountforPayment || 0,

      status: item.Status || "",
    }));

    setData(formatted);
  } catch (error) {
    console.error("Data error:", error);
  }
};
 const getCapexData12 = async () => {
  try {
    debugger;

    let filterQuery = "";

    const currentUser = await sp.web.currentUser();

    // ✅ My Request
    if (activeMenu === "My Request") {
      filterQuery = `Author/EMail eq '${currentUser.Email}' and Status ne 'Paid' and Status ne 'Rejected'`;
    }

    // ✅ Paid Tab
    else if (activeMenu === "Paid") {
      filterQuery = "Status eq 'Paid'";
    }

    // ✅ Rejected Tab
    else if (activeMenu === "Rejected") {
      filterQuery = "Status eq 'Rejected'";
    }

    const items = await sp.web.lists
      .getByTitle("CapexPayment")
      .items.select(
        "ID",
        "Title",
        "Created",
        "EmployeeName",

        // ✅ LOOKUP FIELD
        "VendorName/Id",
        //"VendorName/VendorCode",
        "VendorName/VendorName",

        "PONumber",
        "RequestedAmountforPayment",
        "Status",

        // ✅ Author
        "Author/EMail"
      )
      .expand("Author", "VendorName")
      .filter(filterQuery)
      .orderBy("ID", false)();

    console.log("Items:", items);

    const formatted = items.map((item: any) => ({
      ID: item.ID,

      id: item.Title,

      date: item.Created
        ? new Date(item.Created).toLocaleDateString("en-GB")
        : "",

      EmployeeName: item.EmployeeName || "",

      // ✅ Vendor Details
      vendor: item.VendorName?.VendorName || "",

     // vendorCode: item.VendorName?.VendorCode || "",

      po: item.PONumber || "",

      amount: item.RequestedAmountforPayment || 0,

      status: item.Status || "",
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

  // ✅ Paid
  if (activeMenu === "Paid") {
    menuFilter = item.status?.toLowerCase() === "paid";
  }

  // ✅ Rejected
  else if (activeMenu === "Rejected") {
    menuFilter = item.status?.toLowerCase() === "rejected";
  }

  // ✅ My Request
  else if (activeMenu === "My Request") {
    menuFilter = true;
  }

  return (
    menuFilter &&
    (
      item.id?.toLowerCase().includes(text) ||
      item.vendor?.toLowerCase().includes(text) ||
      item.po?.toLowerCase().includes(text)
    ) &&
    (!status || item.status?.toLowerCase().includes(status))
  );
});


  
  const handleFormOpen = async (item: any, type: "view" | "edit") => {
    try {
      const fullItem = await sp.web.lists
  .getByTitle("CapexPayment")
  .items.getById(item.ID)
  .select(
    "ID",
    "Title",
    "EmployeeName",
    "CapexId",
    "VendorName",
   // "VendorCode",
    "PONumber",
    "PODate",
    "POAmount",
    "POPaymentTerms",              // ✅ ADD
    "MRNNumber",
    "MRNDtae",
    "MRNAmountwithGST",           // ✅ ADD
    "RequestedAmountforPayment",
    "FinalPaymentAgainstPO",      // ✅ ADD
    "InstallationDetails",        // ✅ ADD
    "Status",
    "Author/EMail"
  )
  .expand("Author")();


      setSelectedItem(fullItem);
      setFormType(type); // ✅ dynamic
      setShowForm(true);
    } catch (error) {
      console.error(`${type} error:`, error);
    }
  };

  // ✅ VIEW CLICK
  // const handleViewClick = async (item: any) => {
  //   try {
  //     const fullItem = await sp.web.lists
  //       .getByTitle("CapexPayment")
  //       .items.getById(item.ID)
  //       .select("*", "PICName/Title")
  //       .expand("PICName")();

  //     setSelectedItem(fullItem);
  //     setFormType("view");
  //     setShowForm(true);
  //   } catch (error) {
  //     console.error("View error:", error);
  //   }
  // };
  // const handleEditClick = async (item: any) => {
  //   try {
  //     const fullItem = await sp.web.lists
  //       .getByTitle("CapexPayment")
  //       .items.getById(item.ID)
  //       .select("*", "PICName/Title")
  //       .expand("PICName")();

  //     setSelectedItem(fullItem);
  //     setFormType("Edit");
  //     setShowForm(true);
  //   } catch (error) {
  //     console.error("View error:", error);
  //   }
  // };
  // ✅ LOAD DATA
  React.useEffect(() => {
    debugger;
    if (!context) return;

    void getLoggedInUser();
    void getCapexData(); // 🔥 will run on menu change
  }, [context, activeMenu]);

  // ✅ OPEN VIEW PAGE
  if (showForm) {
    if (formType === "view") {
      return (
        <ViewAdvanceForm
          context={context}
          formData={selectedItem}
          onClose={() => {
            setShowForm(false);
            setFormType(null);
            void getCapexData();
          }}
        />
      );
    }

    if (formType === "new") {
      return (
        <NewAdvanceform
          context={context}
          onClose={() => {
            setShowForm(false);
            setFormType(null);
            void getCapexData();
          }}
        />
      );
    }

    if (formType === "edit") {
      return (
        <EditAdvanceForm
          context={context}
          formData={selectedItem}
          onClose={() => {
            setShowForm(false);
            setFormType(null);
            getCapexData();
          }}
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
              <h2> Capex Payment User Dashboard </h2>
            </div>
          </div>
        </div>
        <div className="mainsecondapprove">
          <div className="mainsecondsmall">
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
          <div>
            <a
              onClick={() => {
                setSelectedItem(null);
                setFormType("new");
                setShowForm(true);
              }}
              className="create-button"
            >
              New Request
            </a>
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
                          <div
                            style={{
                              display: "flex",
                              gap: "10px",
                              alignItems: "center",
                            }}
                          >
                            {/* VIEW */}
                            <span
                              onClick={() => handleFormOpen(item, "view")}
                              style={{ cursor: "pointer" }}
                              title="View"
                            >
                              <img src={View} width={15} alt="View" />
                            </span>

                            {/* EDIT */}
                            {(item.status === "Draft" ||
                              item.status === "Send Back") && (
                              <span
                                onClick={() => handleFormOpen(item, "edit")}
                                style={{ cursor: "pointer" }}
                                title="Edit"
                              >
                                <img src={Edit} width={15} alt="Edit" />
                              </span>
                            )}
                          </div>
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

export default UserDashboard;
