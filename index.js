    const tools = [
      { name: "Excel Formatter - Pepsico", desc: "Format data Excel Pepsico DETAIL", link: "https://tranquil-gnome-f34774.netlify.app/", icon: "📊", category: "excel" },
      { name: "Alokasi Overstok", desc: "Manajemen dan alokasi barang overstok", link: "https://moonlit-gecko-1ab281.netlify.app/", icon: "📦", category: "logistik" },
      { name: "Template Alokasi Barang", desc: "Template alokasi dan manajemen stok gudang", link: "https://2emlmu.netlify.app/", icon: "🏭", category: "logistik" },
      { name: "Cek Pendingan", desc: "Mengecek status pendingan pemenuhan", link: "https://sage-conkies-fd072b.netlify.app/", icon: "⏳", category: "logistik" },
      { name: "Table to Email", desc: "Konversi tabel HTML ke format email", link: "https://tabletoemail.netlify.app/", icon: "✉️", category: "konversi" },
      { name: "PO Direct", desc: "Splitter PO untuk manajemen pesanan", link: "https://podirect.netlify.app/", icon: "✂️", category: "po" },
      { name: "PO Auto Reader", desc: "Baca dan proses file PO otomatis", link: "https://readerpo.netlify.app/", icon: "🤖", category: "po" },
      { name: "Fake Strava", desc: "Generator GPX lari simulasi", link: "https://fakeh3.netlify.app/", icon: "🏃", category: "lainnya" },
      { name: "Kalkulator ED", desc: "Kalkulator ED", link: "https://kalkuede.netlify.app/", icon: "📅", category: "kalkulator" },
      { name: "H3 Store", desc: "Toko online solusi digital", link: "https://newh3store.netlify.app/", icon: "🛒", category: "lainnya" },
      { name: "Cons Dropping", desc: "Analisis dan konsolidasi dropping", link: "https://consdroping.netlify.app/", icon: "🚚", category: "logistik" },
      { name: "Qris Pembayaran", desc: "Generator QRIS pembayaran", link: "https://qrish3.netlify.app/", icon: "💳", category: "lainnya" },
      { name: "Excel Line Converter", desc: "Baris Excel menjadi satu baris", link: "https://line2baru.netlify.app/", icon: "↔️", category: "excel" },
      { name: "Migrasi CSV", desc: "Migrasi dan proses CSV", link: "https://wlcmlmg.netlify.app/", icon: "📁", category: "konversi" },
      { name: "Data Tabular", desc: "Tabel data interaktif", link: "https://dtblu.netlify.app/", icon: "📋", category: "konversi" },
      { name: "Excel to JSON", desc: "Konversi Excel ke JSON", link: "https://excelkejson.netlify.app/", icon: "🔄", category: "konversi" },
      { name: "Excel Multi Column Pivot", desc: "Pivot multi-kolom dinamis", link: "https://excel-multi-column-pivot.netlify.app/", icon: "🔁", category: "excel" },
      { name: "Kalkulator Kubikasi Berat", desc: "Hitung kubikasi berat logistik", link: "https://brtkbik.netlify.app/", icon: "⚖️", category: "kalkulator" },
      { name: "Kalkulator Rasio", desc: "Hitung rasio proporsional", link: "https://rasiokalc.netlify.app/", icon: "➗", category: "kalkulator" },
      { name: "Jadwal Kirim & PL", desc: "Jadwal kirim dan pick-up list", link: "https://jadwalkirimpl.netlify.app/", icon: "🗓️", category: "logistik" },
      { name: "Migrasi PO Principal", desc: "Migrasi PO Principal", link: "https://quiet-speculoos-21e41f.netlify.app/", icon: "📤", category: "po" },
      { name: "PO Splitter V2", desc: "Splitter PO versi terbaru untuk manajemen pesanan", link: "https://pospliterv2.vercel.app/", icon: "✂️", category: "po" }
    ];

    const labels = {
      excel: "Excel",
      po: "PO",
      kalkulator: "Kalkulator",
      konversi: "Konversi",
      logistik: "Logistik",
      lainnya: "Lainnya"
    };

    const grid = document.getElementById("toolsGrid");
    const count = document.getElementById("toolCount");
    const search = document.getElementById("searchInput");
    const menuButton = document.getElementById("menuButton");
    const menuClose = document.getElementById("menuClose");
    const menuBackdrop = document.getElementById("menuBackdrop");
    const sideMenu = document.getElementById("sideMenu");
    const themeButton = document.getElementById("themeButton");
    const menuTheme = document.getElementById("menuTheme");
    const quickTheme = document.getElementById("quickTheme");
    const cartButton = document.getElementById("cartButton");
    const floatPlus = document.getElementById("floatPlus");
    const quickMenu = document.getElementById("quickMenu");
    count.textContent = tools.length;

    function openMenu() {
      menuBackdrop.hidden = false;
      requestAnimationFrame(() => {
        menuBackdrop.classList.add("is-open");
        sideMenu.classList.add("is-open");
        sideMenu.setAttribute("aria-hidden", "false");
      });
    }

    function closeMenu() {
      menuBackdrop.classList.remove("is-open");
      sideMenu.classList.remove("is-open");
      sideMenu.setAttribute("aria-hidden", "true");
      setTimeout(() => {
        menuBackdrop.hidden = true;
      }, 180);
    }

    function toggleTheme() {
      document.body.classList.toggle("is-dark");
      const isDark = document.body.classList.contains("is-dark");
      themeButton.textContent = isDark ? "☀️" : "🌙";
    }

    function scrollToTools() {
      document.getElementById("tools").scrollIntoView({ behavior: "smooth" });
    }

    function addSpark(x, y) {
      for (let i = 0; i < 6; i += 1) {
        const spark = document.createElement("span");
        spark.className = "click-spark";
        const degrees = (360 / 6) * i + Math.random() * 30 - 15;
        const angle = degrees * Math.PI / 180;
        const distance = 20 + Math.random() * 25;
        spark.style.left = `${x}px`;
        spark.style.top = `${y}px`;
        spark.style.setProperty("--tx", `${Math.cos(angle) * distance}px`);
        spark.style.setProperty("--ty", `${Math.sin(angle) * distance}px`);
        document.body.appendChild(spark);
        setTimeout(() => spark.remove(), 460);
      }
    }

    function bindCardEffects() {
      document.querySelectorAll(".tool-card").forEach((card) => {
        card.addEventListener("mousemove", (event) => {
          const rect = card.getBoundingClientRect();
          const x = (event.clientX - rect.left) / rect.width - 0.5;
          const y = (event.clientY - rect.top) / rect.height - 0.5;
          card.style.setProperty("--rx", `${y * -2.5}deg`);
          card.style.setProperty("--ry", `${x * 2.5}deg`);
        });
        card.addEventListener("mouseleave", () => {
          card.style.setProperty("--rx", "0deg");
          card.style.setProperty("--ry", "0deg");
        });
      });
    }

    function render(items) {
      grid.innerHTML = "";
      if (!items.length) {
        grid.innerHTML = '<article class="tool-card"><h3>Tidak ada tools</h3><p>Coba gunakan kata kunci lain.</p></article>';
        return;
      }

      items.forEach((tool, index) => {
        const card = document.createElement("article");
        card.className = "tool-card";
        card.style.animationDelay = `${Math.min(index, 12) * 70}ms`;
        card.innerHTML = `
          <span class="tool-icon">${tool.icon}</span>
          <span class="tool-number">#${String(index + 1).padStart(2, "0")}</span>
          <span class="tool-meta">${labels[tool.category] || "Tools"}</span>
          <h3>${tool.name}</h3>
          <p>${tool.desc}</p>
          <div class="tool-actions">
            <a class="btn btn--dark" href="${tool.link}" target="_blank" rel="noreferrer">Buka project →</a>
            <button class="btn" type="button" data-link="${tool.link}">Copy link</button>
          </div>
        `;
        grid.appendChild(card);
      });
      bindCardEffects();
    }

    function applyFilter() {
      const query = search.value.trim().toLowerCase();
      const filtered = tools.filter((tool) => {
        const matchesSearch =
          tool.name.toLowerCase().includes(query) ||
          tool.desc.toLowerCase().includes(query) ||
          tool.category.toLowerCase().includes(query);
        return matchesSearch;
      });
      render(filtered);
    }

    search.addEventListener("input", applyFilter);

    menuButton.addEventListener("click", openMenu);
    menuClose.addEventListener("click", closeMenu);
    menuBackdrop.addEventListener("click", closeMenu);
    sideMenu.querySelectorAll("a").forEach((link) => {
      link.addEventListener("click", closeMenu);
    });

    themeButton.addEventListener("click", toggleTheme);
    menuTheme.addEventListener("click", toggleTheme);
    quickTheme.addEventListener("click", toggleTheme);
    cartButton.addEventListener("click", scrollToTools);

    floatPlus.addEventListener("click", () => {
      quickMenu.classList.toggle("is-open");
      floatPlus.textContent = quickMenu.classList.contains("is-open") ? "×" : "+";
    });

    document.addEventListener("click", (event) => {
      addSpark(event.clientX, event.clientY);
    });

    grid.addEventListener("click", async (event) => {
      const button = event.target.closest("button[data-link]");
      if (!button) return;
      try {
        await navigator.clipboard.writeText(button.dataset.link);
        button.textContent = "Tersalin";
      } catch {
        button.textContent = "Gagal";
      }
      setTimeout(() => {
        button.textContent = "Copy link";
      }, 1200);
    });

    render(tools);
