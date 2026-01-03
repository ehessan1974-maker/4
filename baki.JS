window.onload = function() {
  alert("من فضلك اختر ملف الإكسل عند أول ضغطة على الصفحة");
  document.body.addEventListener("click", function() {
    document.getElementById("uploadExcel").click();
  }, { once: true });
};

document.getElementById('uploadExcel').addEventListener('change', function(e) {
  const reader = new FileReader();
  reader.onload = function(e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const json = XLSX.utils.sheet_to_json(sheet);

    const اليوم = new Date();
    const تاريخ_اليوم = new Date(اليوم.getFullYear(), اليوم.getMonth(), اليوم.getDate());
    const تاريخ_الغد = new Date(اليوم.getFullYear(), اليوم.getMonth(), اليوم.getDate() + 1);

    function تحويل_إلى_وقت(قيمة) {
      const ساعات_بالوقت = قيمة * 24;
      const ساعات = Math.floor(ساعات_بالوقت);
      const دقائق = Math.round((ساعات_بالوقت - ساعات) * 60);
      return `${ساعات}:${دقائق.toString().padStart(2,'0')}`;
    }

    function وقت_كـDate(قيمة, التاريخ) {
      const ساعات_بالوقت = قيمة * 24;
      const ساعات = Math.floor(ساعات_بالوقت);
      const دقائق = Math.round((ساعات_بالوقت - ساعات) * 60);
      return new Date(التاريخ.getFullYear(), التاريخ.getMonth(), التاريخ.getDate(), ساعات, دقائق);
    }

    function صف_حسب_تاريخ(تاريخ) {
      return json.find(row => {
        const قيمة = row["التاريخ ميلادي"];
        if (typeof قيمة === 'number') {
          const d = XLSX.SSF.parse_date_code(قيمة);
          const تاريخ_الصف = new Date(d.y, d.m - 1, d.d);
          return تاريخ_الصف.getTime() === تاريخ.getTime();
        }
        return false;
      });
    }

    const صف_اليوم = صف_حسب_تاريخ(تاريخ_اليوم);
    const صف_الغد = صف_حسب_تاريخ(تاريخ_الغد);

    if (صف_اليوم) {
      document.getElementById("الفجر").textContent   = تحويل_إلى_وقت(صف_اليوم["الفجر"])   + " ص";
      document.getElementById("الشروق").textContent  = تحويل_إلى_وقت(صف_اليوم["الشروق"])  + " ص";
      document.getElementById("الظهر").textContent   = تحويل_إلى_وقت(صف_اليوم["الظهر"])   + " م";
      document.getElementById("العصر").textContent   = تحويل_إلى_وقت(صف_اليوم["العصر"])   + " م";
      document.getElementById("المغرب").textContent  = تحويل_إلى_وقت(صف_اليوم["المغرب"])  + " م";
      document.getElementById("العشاء").textContent  = تحويل_إلى_وقت(صف_اليوم["العشاء"])  + " م";
    }

    const فترات_الإقامة = {
      "الفجر": 30,
      "الظهر": 30,
      "العصر": 20,
      "المغرب": 10,
      "العشاء": 10
    };

    function صيغة_عد(فرق_ميلي) {
      const فرق = Math.max(0, فرق_ميلي);
      const ساعات = Math.floor(فرق / (1000*60*60));
      const دقائق = Math.floor((فرق % (1000*60*60)) / (1000*60));
      const ثواني = Math.floor((فرق % (1000*60)) / 1000);
      return `${ساعات.toString().padStart(2,'0')}:${دقائق.toString().padStart(2,'0')}:${ثواني.toString().padStart(2,'0')}`;
    }

    function تحديث_العد_التنازلي() {
      const الآن = new Date();

      const مواقيت_اليوم = صف_اليوم ? [
        { اسم: "الفجر",   وقت: وقت_كـDate(صف_اليوم["الفجر"], تاريخ_اليوم) },
        { اسم: "الشروق",  وقت: وقت_كـDate(صف_اليوم["الشروق"], تاريخ_اليوم) },
        { اسم: "الظهر",   وقت: وقت_كـDate(صف_اليوم["الظهر"], تاريخ_اليوم) },
        { اسم: "العصر",   وقت: وقت_كـDate(صف_اليوم["العصر"], تاريخ_اليوم) },
        { اسم: "المغرب",  وقت: وقت_كـDate(صف_اليوم["المغرب"], تاريخ_اليوم) },
        { اسم: "العشاء",  وقت: وقت_كـDate(صف_اليوم["العشاء"], تاريخ_اليوم) }
      ] : [];

      let القادمة_للأذان = null;
      for (let i = 0; i < مواقيت_اليوم.length; i++) {
        const وقت_الأذان = مواقيت_اليوم[i].وقت;
        const فترة = فترات_الإقامة[مواقيت_اليوم[i].اسم] || 0;
        const وقت_الإقامة = new Date(وقت_الأذان.getTime() + فترة * 60000);

        if (الآن < وقت_الأذان) {
          القادمة_للأذان = مواقيت_اليوم[i];
          break;
        }
        if (الآن >= وقت_الأذان && الآن < وقت_الإقامة) {
          const التالي = i + 1 < مواقيت_اليوم.length ? مواقيت_اليوم[i + 1] : null;
          if (التالي) القادمة_للأذان = التالي;
          break;
        }
      }

      let الحالية_للإقامة = null;
      for (let i = 0; i < مواقيت_اليوم.length; i++) {
        const وقت_الأذان = مواقيت_اليوم[i].وقت;
        const فترة = فترات_الإقامة[مواقيت_اليوم[i].اسم] || 0;
        const وقت_الإقامة = new Date(وقت_الأذان.getTime() + فترة * 60000);
        if (الآن >= وقت_الأذان && الآن < وقت_الإقامة) {
          الحالية_للإقامة = { اسم: مواقيت_اليوم[i].اسم, وقت_الإقامة };
          break;
        }
      }

      let فرق_أذان = null;
      let فرق_إقامة = null;

      if (القادمة_للأذان) {
        فرق_أذان = القادمة_للأذان.وقت - الآن;
        document.getElementById("باقي للصلاة").textContent =
          `باقي ${صيغة_عد(فرق_أذان)} حتى صلاة ${القادمة_للأذان.اسم}`;
      } else if (صف_الغد) {
        const وقت_فجر_الغد = وقت_كـDate(صف_الغد["الفجر"], تاريخ_الغد);
        فرق_أذان = وقت_فجر_الغد - الآن;
        document.getElementById("باقي للصلاة").textContent =
          `باقي ${صيغة_عد(فرق_أذان)} حتى صلاة الفجر غداً`;
      } else {
        document.getElementById("باقي للصلاة").textContent = "انتهت صلوات اليوم";
      }

      if (الحالية_للإقامة) {
        فرق_إقامة = الحالية_للإقامة.وقت_الإقامة - الآن;
        document.getElementById("باقي لإقامة الصلاة").textContent =
          `باقي ${صيغة_عد(فرق_إقامة)} حتى إقامة صلاة ${الحالية_للإقامة.اسم}`;
      } else if (القادمة_للأذان) {
        const فترة = فترات_الإقامة[القادمة_للأذان.اسم] || 0;
        const وقت_إقامة_القادمة = new Date(القادمة_للأذان.وقت.getTime() + فترة * 60000);
        فرق_إقامة = وقت_إقامة_القادمة - الآن;
        document.getElementById("باقي لإقامة الصلاة").textContent =
          `باقي ${صيغة_عد(فرق_إقامة)} حتى إقامة صلاة ${القادمة_للأذان.اسم}`;
      } else if (صف_الغد) {
        const وقت_فجر_الغد = وقت_كـDate(صف_الغد["الفجر"], تاريخ_الغد);
        const وقت_إقامة_الغد = new Date(وقت_فجر_الغد.getTime() + (فترات_الإقامة["الفجر"] || 0) * 60000);
        فرق_إقامة = وقت_إقامة_الغد - الآن;
        document.getElementById("باقي لإقامة الصلاة").textContent =
          `باقي ${صيغة_عد(فرق_إقامة)} حتى إقامة صلاة الفجر غداً`;
      } else {
        document.getElementById("باقي لإقامة الصلاة").textContent = "";
      }

      // ✅ تلوين الخانة الأقرب بين العد التنازلي للأذان والإقامة
      const خانة_الأذان = document.getElementById("باقي للصلاة");
      const خانة_الإقامة = document.getElementById("باقي لإقامة الصلاة");

      خانة_الأذان.style.backgroundColor = "";
      خانة_الإقامة.style.backgroundColor = "";

      if (فرق_أذان !== null && فرق_إقامة !== null) {
        if (فرق_أذان <= فرق_إقامة) {
          خانة_الأذان.style.backgroundColor = "lightgreen";
          خانة_الأذان.style.color = "darkred";               
          خانة_الأذان.style.borderTop = "2px solid darkorange";     // الحافة العلوية
          خانة_الأذان.style.borderBottom = "2px solid darkorange";  // الحافة السفلية
        } else {
          خانة_الإقامة.style.backgroundColor = "lightgreen";
          خانة_الإقامة.style.color = "darkred";
          خانة_الإقامة.style.borderTop = "2px solid darkorange";
          خانة_الإقامة.style.borderBottom = "2px solid darkorange";
        }
      } else if (فرق_أذان !== null) {
        خانة_الأذان.style.backgroundColor = "lightgreen";
        خانة_الأذان.style.color = "darkred";
        خانة_الأذان.style.borderTop = "2px solid darkorange";     // الحافة العلوية
        خانة_الأذان.style.borderBottom = "2px solid darkorange";  // الحافة السفلية
      } else if (فرق_إقامة !== null) {
        خانة_الإقامة.style.backgroundColor = "lightgreen";
        خانة_الإقامة.style.color = "darkred";
        خانة_الإقامة.style.borderTop = "2px solid darkorange";
        خانة_الإقامة.style.borderBottom = "2px solid darkorange";      }

      // ✅ تلوين خانة مواقيت الصلاة الأقرب
      const كل_الخانات = ["الفجر","الشروق","الظهر","العصر","المغرب","العشاء"];
      كل_الخانات.forEach(id => {
        document.getElementById(id).style.backgroundColor = "";
      });

      if (الحالية_للإقامة) {
        document.getElementById(الحالية_للإقامة.اسم).style.backgroundColor = "lightgreen";
        document.getElementById(الحالية_للإقامة.اسم).style.color = "darkred";
        document.getElementById(الحالية_للإقامة.اسم).style.borderTop = "8px solid darkorange";
        document.getElementById(الحالية_للإقامة.اسم).style.borderBottom = "8px solid darkorange";
      } else if (القادمة_للأذان) {
        document.getElementById(القادمة_للأذان.اسم).style.backgroundColor = "lightgreen";
        document.getElementById(القادمة_للأذان.اسم).style.color = "darkred";
        document.getElementById(القادمة_للأذان.اسم).style.borderTop = "8px solid darkorange";
        document.getElementById(القادمة_للأذان.اسم).style.borderBottom = "8px solid darkorange"; }
    }

    تحديث_العد_التنازلي();
    setInterval(تحديث_العد_التنازلي, 1000); // تحديث كل ثانية
  };
  reader.readAsArrayBuffer(e.target.files[0]);
});
