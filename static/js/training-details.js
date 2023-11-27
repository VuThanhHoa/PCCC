window.onload = function() {
  var rows = document.querySelectorAll('#training-table tbody tr');
  var allDone = Array.from(rows).every(function(row) {
    return row.classList.contains('completed');
  });

  if (allDone) {
    document.getElementById('training-table').classList.add('completed');
    document.getElementById('new-history').style.display = 'block';
  } else {
    setTimeout(function() {
      location.reload();
    }, 3000);  
  }
};