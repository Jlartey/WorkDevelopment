<script language="javascript">
function formatDateTime(d) {
  var months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  var day = (d.getDate()).toString().padStart(2, '0');
  var mon = months[d.getMonth()];
  var year = d.getFullYear();
  var hours = d.getHours().toString().padStart(2, '0');
  var mins = d.getMinutes().toString().padStart(2, '0');
  var secs = d.getSeconds().toString().padStart(2, '0');
  return day + ' ' + mon + ' ' + year + ' ' + hours + ':' + mins + ':' + secs;
}

function adjustDateTime(fieldId, isEndOfDay) {
  var field = document.getElementById(fieldId);
  if (!field) return;

  // Flag to avoid infinite loops
  var isAdjusting = false;

  var observer = new MutationObserver(function(mutations) {
    if (isAdjusting) return;
    mutations.forEach(function(mutation) {
      if (mutation.type === 'attributes' && mutation.attributeName === 'value') {
        var val = field.value;
        if (val && val.trim() !== '') {
          isAdjusting = true;
          var d = new Date(val);
          if (isNaN(d.getTime())) {
            // If invalid date, try parsing assuming dd mmm yyyy hh:mm:ss format
            var parts = val.match(/(\d{1,2})\s+(\w{3})\s+(\d{4})\s+(\d{2}):(\d{2}):(\d{2})/);
            if (parts) {
              var monthIndex = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'].indexOf(parts[2]);
              d = new Date(parseInt(parts[3]), monthIndex, parseInt(parts[1]), parseInt(parts[4]), parseInt(parts[5]), parseInt(parts[6]));
            }
          }
          if (!isNaN(d.getTime())) {
            if (isEndOfDay) {
              d.setHours(23, 59, 59, 999);
            } else {
              // Keep selected date, but apply current time
              var now = new Date();
              d.setHours(now.getHours(), now.getMinutes(), now.getSeconds(), now.getMilliseconds());
            }
            field.value = formatDateTime(d);
          }
          isAdjusting = false;
        }
      }
    });
  });

  observer.observe(field, { attributes: true });
}

// Apply adjustments after DOM loads
if (document.addEventListener) {
  document.addEventListener('DOMContentLoaded', function() {
    adjustDateTime(usrDate||IM081^^IM081.1Column2', false); // Admission: current time
    adjustDateTime('IM081^^IM081.1Column5', true);   // Discharge: 23:59:59
  });
} else if (document.attachEvent) {
  document.attachEvent('onreadystatechange', function() {
    if (document.readyState === 'complete') {
      adjustDateTime('usrDate||IM081^^IM081.1Column2', false); // Admission: current time
      adjustDateTime('IM081^^IM081.1Column5', true);   // Discharge: 23:59:59
    }
  });
}
</script>
