let students = [];

// Upload Excel
document.getElementById("excelFile").addEventListener("change", function(e) {
  let file = e.target.files[0];
  let reader = new FileReader();
  reader.onload = function(event) {
    let data = new Uint8Array(event.target.result);
    let workbook = XLSX.read(data, {type: "array"});
    let sheet = workbook.Sheets[workbook.SheetNames[0]];
    let rows = XLSX.utils.sheet_to_json(sheet, {header:1});
    students = rows.slice(1).map(r => ({
      id: r[0], name: r[1], class: r[2], father: r[3]
    }));
    renderCards();
  };
  reader.readAsArrayBuffer(file);
});

// Add Manual
function addManual() {
  let id = document.getElementById("id").value;
  let name = document.getElementById("name").value;
  let cls = document.getElementById("class").value;
  let father = document.getElementById("father").value;
  if(id && name) {
    students.push({id, name, class: cls, father});
    renderCards();
    document.getElementById("id").value="";
    document.getElementById("name").value="";
    document.getElementById("class").value="";
    document.getElementById("father").value="";
  }
}

// Render Admit Cards
function renderCards() {
  let container = document.getElementById("cards");
  container.innerHTML = "";
  students.forEach(s => {
    let card = document.createElement("div");
    card.className = "admit-card";
    card.innerHTML = `
      <div class="admit-header">
        <img src="logo.png" alt="logo">
        <h3>NANDI GURUKUL VIDYA MANDIR <br> HALF YEARLY EXAM 2025-26 <br> ADMIT CARD</h3>
      </div>
      <div class="details">
        <p><b>Student Id:</b> ${s.id}</p>
        <p><b>Student Name:</b> ${s.name}</p>
        <p><b>Class:</b> ${s.class}</p>
        <p><b>Father's Name:</b> ${s.father}</p>
      </div>
      <div class="signature">
        <img src="signature.png" alt="sign">
        <p>Principal's Signature</p>
      </div>
    `;
    container.appendChild(card);
  });
}

// Print
function printAdmitCards() {
  window.print();
}
