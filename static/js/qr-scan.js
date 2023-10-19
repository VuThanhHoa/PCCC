let videoStream;
        let frontCamera = false;
        const video = document.getElementById('video');
        const canvas = document.createElement('canvas');
        const context = canvas.getContext('2d');

        async function openCamera() {
            try {
                videoStream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: (frontCamera ? 'user' : 'environment') } });
                video.srcObject = videoStream;
            } catch (err) {
                console.error('Error accessing the camera: ', err);
            }
        }

        function toggleCamera() {
            frontCamera = !frontCamera;
            if (videoStream) {
                videoStream.getTracks().forEach(track => track.stop());
            }
            openCamera();
        }

        function capturePhoto() {
            canvas.width = video.videoWidth;
            canvas.height = video.videoHeight;
            context.drawImage(video, 0, 0, canvas.width, canvas.height);
            // document.getElementById('captureButton').disabled = true;
            // Chuyển ảnh thành base64 để gửi về server
            const imageBase64 = canvas.toDataURL('image/png');
            
            // Gửi ảnh về server
            fetch('/user/qr-scan', {
                method: 'POST',
                body: JSON.stringify({ image: imageBase64 }),
                headers: {
                    'Content-Type': 'application/json'
                }
            })
            .then((response) => response.json())
            .then((data) => {
                // console.error('Error:', data);
                // window.location.href = data.redirect_url;
                document.getElementById('status').innerText = data.status;
                if (/Điểm danh (.*?) thành công/.test(data.status)) {
                    // Lấy tên nhân viên từ phản hồi
                    var employeeName = data.status.match(/Điểm danh (.*?) thành công/)[1];
                    console.log(employeeName);
                    removeTableRowByEmployeeName(employeeName);
                }
            })
            .catch((error) => {
                console.error('Error:', error);
            });
        }
        openCamera();
        function removeTableRowByEmployeeName(employeeName) {
            var table = document.querySelector('table'); // Lấy thẻ table
            var rows = table.querySelectorAll('tr'); // Lấy tất cả các dòng trong bảng

            rows.forEach(function (row) {
                var cell = row.querySelector('td:nth-child(2)'); // Lấy ô thứ hai (Họ tên)

                if (cell && cell.textContent.trim() === employeeName) {
                    row.remove(); // Xóa dòng nếu tên nhân viên trùng khớp
                }
            });
        }
