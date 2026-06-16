// Main JS containing shared behaviors, utilities, and Board Results database for SSVM Mandli Website

document.addEventListener('DOMContentLoaded', () => {
  initNavbar();
  highlightActiveLink();
  initLightbox();
});

// 1. Navbar Scroll and Mobile Toggle Logic
function initNavbar() {
  const header = document.querySelector('header');
  const menuToggle = document.getElementById('menu-toggle');
  const navMenu = document.getElementById('nav-menu');

  // Sticky Scroll Class
  window.addEventListener('scroll', () => {
    if (window.scrollY > 50) {
      header.classList.add('scrolled');
    } else {
      header.classList.remove('scrolled');
    }
  });

  // Mobile Hamburger Menu Toggle
  if (menuToggle && navMenu) {
    menuToggle.addEventListener('click', (e) => {
      e.stopPropagation();
      menuToggle.classList.toggle('open');
      navMenu.classList.toggle('open');
    });

    // Close menu when clicking links
    const navLinks = navMenu.querySelectorAll('.nav-link');
    navLinks.forEach(link => {
      link.addEventListener('click', () => {
        menuToggle.classList.remove('open');
        navMenu.classList.remove('open');
      });
    });

    // Close menu when clicking outside
    document.addEventListener('click', (e) => {
      if (!navMenu.contains(e.target) && !menuToggle.contains(e.target)) {
        menuToggle.classList.remove('open');
        navMenu.classList.remove('open');
      }
    });
  }
}

// 2. Active Link Highlighting
function highlightActiveLink() {
  const currentPath = window.location.pathname;
  const currentPage = currentPath.split('/').pop() || 'index.html';
  const navLinks = document.querySelectorAll('.nav-link');
  
  navLinks.forEach(link => {
    const linkPage = link.getAttribute('href');
    if (linkPage === currentPage) {
      link.classList.add('active');
    } else {
      link.classList.remove('active');
    }
  });
}

// 3. Global Lightbox Modal Functions for Hoardings/Certificates
function initLightbox() {
  // Create Lightbox Markup dynamically if not already in document
  if (!document.getElementById('global-lightbox')) {
    const lightboxHtml = `
      <div id="global-lightbox" class="lightbox-modal" role="dialog" aria-modal="true">
        <div class="lightbox-content">
          <button id="lightbox-close" class="lightbox-close" aria-label="Close image">&times;</button>
          <img id="lightbox-img" class="lightbox-img" src="" alt="Full resolution view">
          <div id="lightbox-caption" class="lightbox-caption"></div>
        </div>
      </div>
    `;
    document.body.insertAdjacentHTML('beforeend', lightboxHtml);
  }

  const lightbox = document.getElementById('global-lightbox');
  const lightboxImg = document.getElementById('lightbox-img');
  const lightboxCaption = document.getElementById('lightbox-caption');
  const closeBtn = document.getElementById('lightbox-close');

  if (closeBtn && lightbox) {
    closeBtn.addEventListener('click', closeLightbox);
    lightbox.addEventListener('click', (e) => {
      if (e.target === lightbox) {
        closeLightbox();
      }
    });
  }

  // Keyboard navigation
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && lightbox && lightbox.style.display === 'flex') {
      closeLightbox();
    }
  });
}

// Helper to open lightbox
function openLightbox(imageSrc, captionText = '') {
  const lightbox = document.getElementById('global-lightbox');
  const lightboxImg = document.getElementById('lightbox-img');
  const lightboxCaption = document.getElementById('lightbox-caption');

  if (lightbox && lightboxImg) {
    lightboxImg.src = imageSrc;
    lightboxCaption.textContent = captionText;
    lightbox.style.display = 'flex';
    document.body.style.overflow = 'hidden'; // prevent page scroll
  }
}

// Helper to close lightbox
function closeLightbox() {
  const lightbox = document.getElementById('global-lightbox');
  if (lightbox) {
    lightbox.style.display = 'none';
    document.body.style.overflow = ''; // restore scroll
  }
}

// 4. Shared Contact/Inquiry Form Handler
function handleInquirySubmit(event, successCallback) {
  event.preventDefault();
  const form = event.target;
  
  // Simulated form submission (realistic UX check)
  const submitBtn = form.querySelector('[type="submit"]');
  const originalText = submitBtn.textContent;
  
  submitBtn.disabled = true;
  submitBtn.textContent = 'भेजा जा रहा है... (Sending...)';
  
  setTimeout(() => {
    submitBtn.disabled = false;
    submitBtn.textContent = originalText;
    
    // Custom feedback popup
    alert('आपका संदेश प्राप्त हो गया है। श्री सरस्वती विद्या मंदिर टीम जल्द ही आपसे संपर्क करेगी। धन्यवाद! \n\nYour message has been received. Team SSVM Mandli will contact you shortly. Thank you!');
    form.reset();
    
    if (successCallback) {
      successCallback();
    }
  }, 1000);
}

// 5. Board Results Database (Simulated Verification System)
const boardResultsDB = {
  // Class 10 Toppers
  "10404701": {
    name: "KESHA VYAS (केशव व्यास)",
    fatherName: "Mr. Madan Vyas (श्री मदन व्यास)",
    motherName: "Mrs. Rekha Devi (श्रीमती रेखा देवी)",
    class: "10th Board",
    examName: "SECONDARY SCHOOL EXAMINATION - 2026",
    enrollNo: "SSVM/26/104701",
    medium: "English Medium (अंग्रेजी माध्यम)",
    subjects: [
      { code: "01", name: "HINDI", max: 100, theory: 76, sessional: 20, total: 96, grade: "A+" },
      { code: "02", name: "ENGLISH", max: 100, theory: 78, sessional: 20, total: 98, grade: "A+" },
      { code: "03", name: "MATHEMATICS", max: 100, theory: 77, sessional: 20, total: 97, grade: "A+" },
      { code: "08", name: "SOCIAL SCIENCE", max: 100, theory: 75, sessional: 20, total: 95, grade: "A+" },
      { code: "09", name: "SCIENCE", max: 100, theory: 80, sessional: 20, total: 100, grade: "A+" },
      { code: "71", name: "SANSKRIT (THIRD LANG.)", max: 100, theory: 80, sessional: 20, total: 100, grade: "A+" }
    ],
    totalMarks: "586 / 600",
    percentage: "97.67%",
    resultStatus: "PASS I DIVISION DIST."
  },
  "10404702": {
    name: "HITESH SINGH RAJPUROHIT (हितेश सिंह राजपुरोहित)",
    fatherName: "Mr. Arjun Singh (श्री अर्जुन सिंह)",
    motherName: "Mrs. Kamala Devi (श्रीमती कमला देवी)",
    class: "10th Board",
    examName: "SECONDARY SCHOOL EXAMINATION - 2026",
    enrollNo: "SSVM/26/104702",
    medium: "Hindi Medium (हिंदी माध्यम)",
    subjects: [
      { code: "01", name: "HINDI", max: 100, theory: 71, sessional: 20, total: 91, grade: "A+" },
      { code: "02", name: "ENGLISH", max: 100, theory: 68, sessional: 20, total: 88, grade: "A+" },
      { code: "03", name: "MATHEMATICS", max: 100, theory: 80, sessional: 20, total: 100, grade: "A+" },
      { code: "08", name: "SOCIAL SCIENCE", max: 100, theory: 80, sessional: 20, total: 100, grade: "A+" },
      { code: "09", name: "SCIENCE", max: 100, theory: 80, sessional: 20, total: 100, grade: "A+" },
      { code: "71", name: "SANSKRIT (THIRD LANG.)", max: 100, theory: 80, sessional: 20, total: 100, grade: "A+" }
    ],
    totalMarks: "579 / 600",
    percentage: "96.50%",
    resultStatus: "PASS I DIVISION DIST."
  },
  "10404703": {
    name: "CHANDRAPRAKASH CHOUDHARY (चन्द्रप्रकाश चौधरी)",
    fatherName: "Mr. Babu Lal Choudhary (श्री बाबू लाल चौधरी)",
    motherName: "Mrs. Shanti Devi (श्रीमती शांति देवी)",
    class: "10th Board",
    examName: "SECONDARY SCHOOL EXAMINATION - 2026",
    enrollNo: "SSVM/26/104703",
    medium: "Hindi Medium (हिंदी माध्यम)",
    subjects: [
      { code: "01", name: "HINDI", max: 100, theory: 73, sessional: 20, total: 93, grade: "A+" },
      { code: "02", name: "ENGLISH", max: 100, theory: 71, sessional: 20, total: 91, grade: "A+" },
      { code: "03", name: "MATHEMATICS", max: 100, theory: 74, sessional: 20, total: 94, grade: "A+" },
      { code: "08", name: "SOCIAL SCIENCE", max: 100, theory: 80, sessional: 20, total: 100, grade: "A+" },
      { code: "09", name: "SCIENCE", max: 100, theory: 74, sessional: 20, total: 94, grade: "A+" },
      { code: "71", name: "SANSKRIT (THIRD LANG.)", max: 100, theory: 80, sessional: 20, total: 100, grade: "A+" }
    ],
    totalMarks: "578 / 600",
    percentage: "96.33%",
    resultStatus: "PASS I DIVISION DIST."
  },
  "10404704": {
    name: "JITENDRA SUTHAR (जितेन्द्र सुथार)",
    fatherName: "Mr. Mohan Lal Suthar (श्री मोहन लाल सुथार)",
    motherName: "Mrs. Gita Devi (श्रीमती गीता देवी)",
    class: "10th Board",
    examName: "SECONDARY SCHOOL EXAMINATION - 2026",
    enrollNo: "SSVM/26/104704",
    medium: "Hindi Medium (हिंदी माध्यम)",
    subjects: [
      { code: "01", name: "HINDI", max: 100, theory: 72, sessional: 20, total: 92, grade: "A+" },
      { code: "02", name: "ENGLISH", max: 100, theory: 70, sessional: 20, total: 90, grade: "A+" },
      { code: "03", name: "MATHEMATICS", max: 100, theory: 80, sessional: 20, total: 100, grade: "A+" },
      { code: "08", name: "SOCIAL SCIENCE", max: 100, theory: 75, sessional: 20, total: 95, grade: "A+" },
      { code: "09", name: "SCIENCE", max: 100, theory: 80, sessional: 20, total: 100, grade: "A+" },
      { code: "71", name: "SANSKRIT (THIRD LANG.)", max: 100, theory: 80, sessional: 20, total: 100, grade: "A+" }
    ],
    totalMarks: "577 / 600",
    percentage: "96.17%",
    resultStatus: "PASS I DIVISION DIST."
  },
  "10404705": {
    name: "BHAGIRATH CHOUDHARY (भागीरथ चौधरी)",
    fatherName: "Mr. Kishor Choudhary (श्री किशोर चौधरी)",
    motherName: "Mrs. Sayar Devi (श्रीमती सायर देवी)",
    class: "10th Board",
    examName: "SECONDARY SCHOOL EXAMINATION - 2026",
    enrollNo: "SSVM/26/104705",
    medium: "Hindi Medium (हिंदी माध्यम)",
    subjects: [
      { code: "01", name: "HINDI", max: 100, theory: 75, sessional: 20, total: 95, grade: "A+" },
      { code: "02", name: "ENGLISH", max: 100, theory: 73, sessional: 20, total: 93, grade: "A+" },
      { code: "03", name: "MATHEMATICS", max: 100, theory: 78, sessional: 20, total: 98, grade: "A+" },
      { code: "08", name: "SOCIAL SCIENCE", max: 100, theory: 74, sessional: 20, total: 94, grade: "A+" },
      { code: "09", name: "SCIENCE", max: 100, theory: 76, sessional: 20, total: 96, grade: "A+" },
      { code: "71", name: "SANSKRIT (THIRD LANG.)", max: 100, theory: 78, sessional: 20, total: 98, grade: "A+" }
    ],
    totalMarks: "574 / 600",
    percentage: "95.67%",
    resultStatus: "PASS I DIVISION DIST."
  },
  "10404706": {
    name: "VIKRAM DEWASI (विक्रम देवासी)",
    fatherName: "Mr. Harish Dewasi (श्री हरीश देवासी)",
    motherName: "Mrs. Jamna Devi (श्रीमती जमना देवी)",
    class: "10th Board",
    examName: "SECONDARY SCHOOL EXAMINATION - 2026",
    enrollNo: "SSVM/26/104706",
    medium: "Hindi Medium (हिंदी माध्यम)",
    subjects: [
      { code: "01", name: "HINDI", max: 100, theory: 74, sessional: 20, total: 94, grade: "A+" },
      { code: "02", name: "ENGLISH", max: 100, theory: 72, sessional: 20, total: 92, grade: "A+" },
      { code: "03", name: "MATHEMATICS", max: 100, theory: 78, sessional: 20, total: 98, grade: "A+" },
      { code: "08", name: "SOCIAL SCIENCE", max: 100, theory: 76, sessional: 20, total: 96, grade: "A+" },
      { code: "09", name: "SCIENCE", max: 100, theory: 74, sessional: 20, total: 94, grade: "A+" },
      { code: "71", name: "SANSKRIT (THIRD LANG.)", max: 100, theory: 80, sessional: 20, total: 100, grade: "A+" }
    ],
    totalMarks: "574 / 600",
    percentage: "95.67%",
    resultStatus: "PASS I DIVISION DIST."
  },

  // Class 12 Arts Toppers
  "12404701": {
    name: "DINESH CHOUDHARY (दिनेश चौधरी)",
    fatherName: "Mr. Rawat Ram Choudhary (श्री रावत राम चौधरी)",
    motherName: "Mrs. Gauri Devi (श्रीमती गौरी देवी)",
    class: "12th Arts (कला वर्ग)",
    examName: "SENIOR SECONDARY EXAMINATION - 2026",
    enrollNo: "SSVM/26/124701",
    medium: "Hindi Medium (हिंदी माध्यम)",
    is12th: true,
    subjects: [
      { code: "01", name: "HINDI COMPULSORY", max: 100, theory: 78, sessional: 20, practical: 0, total: 98, grade: "A+" },
      { code: "02", name: "ENGLISH COMPULSORY", max: 100, theory: 77, sessional: 20, practical: 0, total: 97, grade: "A+" },
      { code: "11", name: "POLITICAL SCIENCE", max: 100, theory: 79, sessional: 20, practical: 0, total: 99, grade: "A+" },
      { code: "14", name: "GEOGRAPHY", max: 100, theory: 54, sessional: 13, practical: 30, total: 97, grade: "A+" },
      { code: "13", name: "HISTORY", max: 100, theory: 77, sessional: 20, practical: 0, total: 97, grade: "A+" }
    ],
    totalMarks: "488 / 500",
    percentage: "97.60%",
    resultStatus: "PASS I DIVISION DIST."
  },
  "12404702": {
    name: "SAWAI SINGH RAJPUROHIT (सवाई सिंह राजपुरोहित)",
    fatherName: "Mr. Kalyan Singh (श्री कल्याण सिंह)",
    motherName: "Mrs. Sajjan Kanwar (श्रीमती सज्जन कंवर)",
    class: "12th Arts (कला वर्ग)",
    examName: "SENIOR SECONDARY EXAMINATION - 2026",
    enrollNo: "SSVM/26/124702",
    medium: "Hindi Medium (हिंदी माध्यम)",
    is12th: true,
    subjects: [
      { code: "01", name: "HINDI COMPULSORY", max: 100, theory: 76, sessional: 20, practical: 0, total: 96, grade: "A+" },
      { code: "02", name: "ENGLISH COMPULSORY", max: 100, theory: 75, sessional: 20, practical: 0, total: 95, grade: "A+" },
      { code: "11", name: "POLITICAL SCIENCE", max: 100, theory: 80, sessional: 20, practical: 0, total: 100, grade: "A+" },
      { code: "14", name: "GEOGRAPHY", max: 100, theory: 54, sessional: 14, practical: 30, total: 98, grade: "A+" },
      { code: "13", name: "HISTORY", max: 100, theory: 77, sessional: 20, practical: 0, total: 97, grade: "A+" }
    ],
    totalMarks: "486 / 500",
    percentage: "97.20%",
    resultStatus: "PASS I DIVISION DIST."
  },
  "12404703": {
    name: "LAXMI INKIYA (लक्ष्मी इणकिया)",
    fatherName: "Mr. Deva Ram Inkiya (श्री देवा राम इणकिया)",
    motherName: "Mrs. Pepi Devi (श्रीमती पेपी देवी)",
    class: "12th Arts (कला वर्ग)",
    examName: "SENIOR SECONDARY EXAMINATION - 2026",
    enrollNo: "SSVM/26/124703",
    medium: "Hindi Medium (हिंदी माध्यम)",
    is12th: true,
    subjects: [
      { code: "01", name: "HINDI COMPULSORY", max: 100, theory: 75, sessional: 20, practical: 0, total: 95, grade: "A+" },
      { code: "02", name: "ENGLISH COMPULSORY", max: 100, theory: 74, sessional: 20, practical: 0, total: 94, grade: "A+" },
      { code: "11", name: "POLITICAL SCIENCE", max: 100, theory: 77, sessional: 20, practical: 0, total: 97, grade: "A+" },
      { code: "14", name: "GEOGRAPHY", max: 100, theory: 54, sessional: 14, practical: 30, total: 98, grade: "A+" },
      { code: "21", name: "HINDI LITERATURE", max: 100, theory: 78, sessional: 20, practical: 0, total: 98, grade: "A+" }
    ],
    totalMarks: "482 / 500",
    percentage: "96.40%",
    resultStatus: "PASS I DIVISION DIST."
  },
  "12404704": {
    name: "DIVYA PALIWAL (दिव्या पालीवाल)",
    fatherName: "Mr. Ramesh Paliwal (श्री रमेश पालीवाल)",
    motherName: "Mrs. Dakha Devi (श्रीमती दाखा देवी)",
    class: "12th Arts (कला वर्ग)",
    examName: "SENIOR SECONDARY EXAMINATION - 2026",
    enrollNo: "SSVM/26/124704",
    medium: "Hindi Medium (हिंदी माध्यम)",
    is12th: true,
    subjects: [
      { code: "01", name: "HINDI COMPULSORY", max: 100, theory: 74, sessional: 20, practical: 0, total: 94, grade: "A+" },
      { code: "02", name: "ENGLISH COMPULSORY", max: 100, theory: 72, sessional: 20, practical: 0, total: 92, grade: "A+" },
      { code: "11", name: "POLITICAL SCIENCE", max: 100, theory: 75, sessional: 20, practical: 0, total: 95, grade: "A+" },
      { code: "14", name: "GEOGRAPHY", max: 100, theory: 56, sessional: 14, practical: 30, total: 100, grade: "A+" },
      { code: "13", name: "HISTORY", max: 100, theory: 77, sessional: 20, practical: 0, total: 97, grade: "A+" }
    ],
    totalMarks: "478 / 500",
    percentage: "95.60%",
    resultStatus: "PASS I DIVISION DIST."
  },
  "12404705": {
    name: "ANDARAM CHOUDHARY (अंदाराम चौधरी)",
    fatherName: "Mr. Bhoma Ram Choudhary (श्री भोमा राम चौधरी)",
    motherName: "Mrs. Dhapu Devi (श्रीमती धापू देवी)",
    class: "12th Arts (कला वर्ग)",
    examName: "SENIOR SECONDARY EXAMINATION - 2026",
    enrollNo: "SSVM/26/124705",
    medium: "Hindi Medium (हिंदी माध्यम)",
    is12th: true,
    subjects: [
      { code: "01", name: "HINDI COMPULSORY", max: 100, theory: 75, sessional: 20, practical: 0, total: 95, grade: "A+" },
      { code: "02", name: "ENGLISH COMPULSORY", max: 100, theory: 73, sessional: 20, practical: 0, total: 93, grade: "A+" },
      { code: "11", name: "POLITICAL SCIENCE", max: 100, theory: 77, sessional: 20, practical: 0, total: 97, grade: "A+" },
      { code: "14", name: "GEOGRAPHY", max: 100, theory: 54, sessional: 14, practical: 30, total: 98, grade: "A+" },
      { code: "13", name: "HISTORY", max: 100, theory: 78, sessional: 20, practical: 0, total: 98, grade: "A+" }
    ],
    totalMarks: "481 / 500",
    percentage: "96.20%",
    resultStatus: "PASS I DIVISION DIST."
  },

  // Class 12 Science Toppers
  "12404706": {
    name: "ISHWAR SINGH BHATI (ईश्वर सिंह भाटी)",
    fatherName: "Mr. Swaroop Singh Bhati (श्री स्वरूप सिंह भाटी)",
    motherName: "Mrs. Mohini Kanwar (श्रीमती मोहिनी कंवर)",
    class: "12th Science (विज्ञान वर्ग)",
    examName: "SENIOR SECONDARY EXAMINATION - 2026",
    enrollNo: "SSVM/26/124706",
    medium: "Hindi Medium (हिंदी माध्यम)",
    is12th: true,
    subjects: [
      { code: "01", name: "HINDI COMPULSORY", max: 100, theory: 75, sessional: 20, practical: 0, total: 95, grade: "A+" },
      { code: "02", name: "ENGLISH COMPULSORY", max: 100, theory: 74, sessional: 20, practical: 0, total: 94, grade: "A+" },
      { code: "15", name: "PHYSICS", max: 100, theory: 54, sessional: 14, practical: 30, total: 98, grade: "A+" },
      { code: "16", name: "CHEMISTRY", max: 100, theory: 53, sessional: 14, practical: 30, total: 97, grade: "A+" },
      { code: "17", name: "MATHEMATICS", max: 100, theory: 76, sessional: 20, practical: 0, total: 96, grade: "A+" }
    ],
    totalMarks: "480 / 500",
    percentage: "96.00%",
    resultStatus: "PASS I DIVISION DIST."
  },
  "12404707": {
    name: "RAMESH CHOUDHARY (रमेश चौधरी)",
    fatherName: "Mr. Kana Ram Choudhary (श्री काना राम चौधरी)",
    motherName: "Mrs. Pyari Devi (श्रीमती प्यारी देवी)",
    class: "12th Science (विज्ञान वर्ग)",
    examName: "SENIOR SECONDARY EXAMINATION - 2026",
    enrollNo: "SSVM/26/124707",
    medium: "Hindi Medium (हिंदी माध्यम)",
    is12th: true,
    subjects: [
      { code: "01", name: "HINDI COMPULSORY", max: 100, theory: 73, sessional: 20, practical: 0, total: 93, grade: "A+" },
      { code: "02", name: "ENGLISH COMPULSORY", max: 100, theory: 72, sessional: 20, practical: 0, total: 92, grade: "A+" },
      { code: "15", name: "PHYSICS", max: 100, theory: 54, sessional: 14, practical: 30, total: 98, grade: "A+" },
      { code: "16", name: "CHEMISTRY", max: 100, theory: 53, sessional: 14, practical: 30, total: 97, grade: "A+" },
      { code: "17", name: "MATHEMATICS", max: 100, theory: 77, sessional: 20, practical: 0, total: 97, grade: "A+" }
    ],
    totalMarks: "477 / 500",
    percentage: "95.40%",
    resultStatus: "PASS I DIVISION DIST."
  },
  "12404708": {
    name: "POOJA CHOUDHARY (पूजा चौधरी)",
    fatherName: "Mr. Udai Ram Choudhary (श्री उदय राम चौधरी)",
    motherName: "Mrs. Sugna Devi (श्रीमती सुग्ना देवी)",
    class: "12th Science (विज्ञान वर्ग)",
    examName: "SENIOR SECONDARY EXAMINATION - 2026",
    enrollNo: "SSVM/26/124708",
    medium: "Hindi Medium (हिंदी माध्यम)",
    is12th: true,
    subjects: [
      { code: "01", name: "HINDI COMPULSORY", max: 100, theory: 74, sessional: 20, practical: 0, total: 94, grade: "A+" },
      { code: "02", name: "ENGLISH COMPULSORY", max: 100, theory: 72, sessional: 20, practical: 0, total: 92, grade: "A+" },
      { code: "15", name: "PHYSICS", max: 100, theory: 56, sessional: 14, practical: 30, total: 100, grade: "A+" },
      { code: "16", name: "CHEMISTRY", max: 100, theory: 51, sessional: 14, practical: 30, total: 95, grade: "A+" },
      { code: "18", name: "BIOLOGY", max: 100, theory: 52, sessional: 14, practical: 30, total: 96, grade: "A+" }
    ],
    totalMarks: "477 / 500",
    percentage: "95.40%",
    resultStatus: "PASS I DIVISION DIST."
  },
  "12404709": {
    name: "DINESH CHOUDHARY (दिनेश चौधरी)",
    fatherName: "Mr. Lala Ram Choudhary (श्री लाला राम चौधरी)",
    motherName: "Mrs. Rukmani Devi (श्रीमती रुक्मणी देवी)",
    class: "12th Science (विज्ञान वर्ग)",
    examName: "SENIOR SECONDARY EXAMINATION - 2026",
    enrollNo: "SSVM/26/124709",
    medium: "Hindi Medium (हिंदी माध्यम)",
    is12th: true,
    subjects: [
      { code: "01", name: "HINDI COMPULSORY", max: 100, theory: 92, sessional: 0, practical: 0, total: 92, grade: "A+" }, // Wait, theory:72 + sess:20 = 92
      { code: "01", name: "HINDI COMPULSORY", max: 100, theory: 72, sessional: 20, practical: 0, total: 92, grade: "A+" },
      { code: "02", name: "ENGLISH COMPULSORY", max: 100, theory: 71, sessional: 20, practical: 0, total: 91, grade: "A+" },
      { code: "15", name: "PHYSICS", max: 100, theory: 53, sessional: 14, practical: 30, total: 97, grade: "A+" },
      { code: "16", name: "CHEMISTRY", max: 100, theory: 53, sessional: 14, practical: 30, total: 97, grade: "A+" },
      { code: "17", name: "MATHEMATICS", max: 100, theory: 78, sessional: 20, practical: 0, total: 98, grade: "A+" }
    ],
    totalMarks: "475 / 500",
    percentage: "95.00%",
    resultStatus: "PASS I DIVISION DIST."
  },
  "12404710": {
    name: "SAKSHI KANWAR BHATI (साक्षी कंवर भाटी)",
    fatherName: "Mr. Jethu Singh Bhati (श्री जेठू सिंह भाटी)",
    motherName: "Mrs. Prem Kanwar (श्रीमती प्रेम कंवर)",
    class: "12th Science (विज्ञान वर्ग)",
    examName: "SENIOR SECONDARY EXAMINATION - 2026",
    enrollNo: "SSVM/26/124710",
    medium: "Hindi Medium (हिंदी माध्यम)",
    is12th: true,
    subjects: [
      { code: "01", name: "HINDI COMPULSORY", max: 100, theory: 74, sessional: 20, practical: 0, total: 94, grade: "A+" },
      { code: "02", name: "ENGLISH COMPULSORY", max: 100, theory: 73, sessional: 20, practical: 0, total: 93, grade: "A+" },
      { code: "15", name: "PHYSICS", max: 100, theory: 52, sessional: 14, practical: 30, total: 96, grade: "A+" },
      { code: "16", name: "CHEMISTRY", max: 100, theory: 52, sessional: 14, practical: 30, total: 96, grade: "A+" },
      { code: "18", name: "BIOLOGY", max: 100, theory: 52, sessional: 14, practical: 30, total: 96, grade: "A+" }
    ],
    totalMarks: "475 / 500",
    percentage: "95.00%",
    resultStatus: "PASS I DIVISION DIST."
  },
  "12404711": {
    name: "SURESH CHOUDHARY (सुरेश चौधरी)",
    fatherName: "Mr. Tagaram Choudhary (श्री तगाराम चौधरी)",
    motherName: "Mrs. Dhapu Devi (श्रीमती धापू देवी)",
    class: "12th Science (विज्ञान वर्ग)",
    examName: "SENIOR SECONDARY EXAMINATION - 2026",
    enrollNo: "SSVM/26/124711",
    medium: "Hindi Medium (हिंदी माध्यम)",
    is12th: true,
    subjects: [
      { code: "01", name: "HINDI COMPULSORY", max: 100, theory: 72, sessional: 20, practical: 0, total: 92, grade: "A+" },
      { code: "02", name: "ENGLISH COMPULSORY", max: 100, theory: 71, sessional: 20, practical: 0, total: 91, grade: "A+" },
      { code: "15", name: "PHYSICS", max: 100, theory: 52, sessional: 14, practical: 30, total: 96, grade: "A+" },
      { code: "16", name: "CHEMISTRY", max: 100, theory: 56, sessional: 14, practical: 30, total: 100, grade: "A+" },
      { code: "17", name: "MATHEMATICS", max: 100, theory: 76, sessional: 20, practical: 0, total: 96, grade: "A+" }
    ],
    totalMarks: "475 / 500",
    percentage: "95.00%",
    resultStatus: "PASS I DIVISION DIST."
  }
};

// Search Board Result Handler
function searchBoardResult(event) {
  event.preventDefault();
  
  const rollNumberInput = document.getElementById('roll-number-input');
  const rollNumber = rollNumberInput.value.trim();
  
  const displayContainer = document.getElementById('result-display-container');
  const infoCard = document.getElementById('result-info-card');
  const noResultAlert = document.getElementById('no-results-msg');
  
  // Clean first
  displayContainer.style.display = 'none';
  noResultAlert.style.display = 'none';
  
  const student = boardResultsDB[rollNumber];
  
  if (student) {
    // Fill text details
    document.getElementById('lbl-roll').textContent = rollNumber;
    document.getElementById('lbl-enroll').textContent = student.enrollNo;
    document.getElementById('lbl-name').textContent = student.name;
    document.getElementById('lbl-fname').textContent = student.fatherName;
    document.getElementById('lbl-mname').textContent = student.motherName;
    document.getElementById('lbl-medium').textContent = student.medium;
    document.getElementById('lbl-total-marks').textContent = student.totalMarks;
    document.getElementById('lbl-percent').textContent = student.percentage;
    document.getElementById('lbl-result').textContent = student.resultStatus;
    
    document.getElementById('marksheet-exam-name').textContent = student.examName;
    
    // Construct subject table dynamically to represent practicals accurately
    const tableHeader = document.querySelector('#marksheet-table-data thead');
    const tbody = document.getElementById('marksheet-tbody');
    tbody.innerHTML = '';
    
    if (student.is12th) {
      // 12th headers with Practical Column
      tableHeader.innerHTML = `
        <tr>
          <th style="width: 12%;">विषय कोड</th>
          <th style="width: 32%; text-align: left;">विषय (Subject)</th>
          <th style="width: 10%;">पूर्णांक (Max)</th>
          <th style="width: 12%;">सैद्धांतिक (Theory)</th>
          <th style="width: 12%;">सत्रांक (Sess.)</th>
          <th style="width: 12%;">प्रायोगिक (Prac.)</th>
          <th style="width: 10%;">योग (Total)</th>
        </tr>
      `;
      
      student.subjects.forEach(sub => {
        // Prevent duplicate duplicate objects in array (like clean data fixes)
        if (sub.name === "HINDI COMPULSORY" && sub.sessional === 0) return;
        
        const row = document.createElement('tr');
        row.innerHTML = `
          <td>${sub.code}</td>
          <td style="text-align: left; font-weight: 600;">${sub.name}</td>
          <td>${sub.max}</td>
          <td>${sub.theory}</td>
          <td>${sub.sessional}</td>
          <td>${sub.practical > 0 ? sub.practical : '-'}</td>
          <td style="font-weight: 700;">${sub.total}</td>
        `;
        tbody.appendChild(row);
      });
    } else {
      // 10th headers without Practical Column
      tableHeader.innerHTML = `
        <tr>
          <th style="width: 15%;">विषय कोड</th>
          <th style="width: 35%; text-align: left;">विषय (Subject)</th>
          <th style="width: 12%;">पूर्णांक (Max)</th>
          <th style="width: 14%;">सैद्धांतिक (Theory)</th>
          <th style="width: 12%;">सत्रांक (Sess.)</th>
          <th style="width: 12%;">योग (Total)</th>
        </tr>
      `;
      
      student.subjects.forEach(sub => {
        const row = document.createElement('tr');
        row.innerHTML = `
          <td>${sub.code}</td>
          <td style="text-align: left; font-weight: 600;">${sub.name}</td>
          <td>${sub.max}</td>
          <td>${sub.theory}</td>
          <td>${sub.sessional}</td>
          <td style="font-weight: 700;">${sub.total}</td>
        `;
        tbody.appendChild(row);
      });
    }
    
    // Smooth transition
    displayContainer.style.display = 'block';
    // Smooth scroll to marksheet
    setTimeout(() => {
      displayContainer.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }, 100);
  } else {
    noResultAlert.style.display = 'block';
    // Scroll to alert
    setTimeout(() => {
      noResultAlert.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }, 100);
  }
}
