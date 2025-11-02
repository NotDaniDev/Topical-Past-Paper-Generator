// Auto-refresh task status every 5 seconds
function pollTaskStatus(taskId) {
  const statusBox = document.getElementById("status-box");
  const progressBox = document.getElementById("progress-box");
  const filesBox = document.getElementById("files-box");

  async function fetchStatus() {
    try {
      const response = await fetch(`/status_api/${taskId}`);
      if (!response.ok) throw new Error("Network response was not ok");
      const data = await response.json();

      // Update status and progress
      statusBox.textContent = data.status;
      progressBox.textContent = data.progress;

      // Update file links if available
      if (data.files) {
        filesBox.innerHTML = "";
        if (data.files.qp) {
          filesBox.innerHTML += `<a class="btn btn-success m-1" href="/download/${taskId}/merged_qp">Download Merged QP</a>`;
        }
        if (data.files.ms) {
          filesBox.innerHTML += `<a class="btn btn-success m-1" href="/download/${taskId}/merged_ms">Download Merged MS</a>`;
        }
        if (data.files.topical) {
          filesBox.innerHTML += "<h5>Topical Files</h5><ul>";
          data.files.topical.forEach((file, idx) => {
            filesBox.innerHTML += `<li><a href="/download/${taskId}/topical_${idx}">${file.topic}</a></li>`;
          });
          filesBox.innerHTML += "</ul>";
        }
      }

      // Keep polling until finished
      if (data.status !== "Completed" && data.status !== "Error") {
        setTimeout(fetchStatus, 5000);
      }
    } catch (err) {
      console.error("Error fetching status:", err);
      progressBox.textContent = "Error fetching status. Retrying...";
      setTimeout(fetchStatus, 10000);
    }
  }

  fetchStatus();
}
