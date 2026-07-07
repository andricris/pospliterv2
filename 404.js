    var el = document.getElementById("path");
    if (el) el.textContent = window.location.pathname || "/";

    // 404 mengikuti arah mouse (tilt parallax)
    var code = document.getElementById("code");
    var raf = null;
    function onMove(e) {
      if (raf) return;
      raf = requestAnimationFrame(function () {
        raf = null;
        var x = e.clientX / window.innerWidth - 0.5;   // -0.5 .. 0.5
        var y = e.clientY / window.innerHeight - 0.5;
        var ry = x * 26;        // putar kiri/kanan
        var rx = 6 - y * 22;    // miring atas/bawah (default 6deg)
        code.style.setProperty("--ry", ry.toFixed(2) + "deg");
        code.style.setProperty("--rx", rx.toFixed(2) + "deg");
      });
    }
    function onLeave() {
      code.style.setProperty("--ry", "-7deg");
      code.style.setProperty("--rx", "6deg");
    }
    if (window.matchMedia("(pointer: fine)").matches) {
      window.addEventListener("mousemove", onMove);
      window.addEventListener("mouseleave", onLeave);
    }

    var backBtn = document.getElementById('backBtn');
    if (backBtn) backBtn.addEventListener('click', function () { history.back(); });
