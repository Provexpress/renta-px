const USER_ACCESS = {
  "gerencia@provexpress.com.co": {
    role: "gerencia",
    comercial: null
  },
  "oscar.perez@provexpress.com.co": {
    role: "gerencia",
    comercial: null
  },
  "especialista.preventa@provexpress.com.co": {
    role: "gerencia",
    comercial: null
  },
  "juannovoa@provexpress.com.co": {
    role: "gerencia",
    comercial: null
  },
   "c.estrategica@provexpress.com.co": {
    role: "gerencia",
    comercial: null
  },
  "finanzas@provexpress.com.co": {
    role: "finanzas",
    comercial: null
  },
  "operaciones@provexpress.com.co": {
    role: "operaciones",
    comercial: null
  },
  "johana.mojica@provexpress.com.co": {
    role: "comercial",
    comercial: "Jasbleidy Mojica",
    comercialAliases: ["Jasbleidy Mojica", "Jasbleidy Johana Mojica Munoz", "Johana Mojica"]
  },
  "astrid.jimenez@provexpress.com.co": {
    role: "comercial",
    comercial: "Astrid Jimenez",
    comercialAliases: ["Astrid Jimenez", "Leidy Astrid Jimenez Ossa"],
    subRenta: true
  },
  "oscar.beltran@provexpress.com.co": {
    role: "comercial",
    comercial: "Oscar Beltran",
    comercialAliases: ["Oscar Beltran", "Oscar Alejandro Beltran Garzon"]
  },
  "lington.linares@provexpress.com.co": {
    role: "comercial",
    comercial: "Lington Linares",
    comercialAliases: ["Lington Linares", "Lington Linares Linares"]
  },
  "angela.torres@provexpress.com.co": {
    role: "comercial",
    comercial: "Angela Torres",
    comercialAliases: ["Angela Torres", "Angela Rocio Torres Matallana"],
    subRenta: true
  },
  "yeison.urrego@provexpress.com.co": {
    role: "comercial",
    comercial: "Yeison Urrego",
    comercialAliases: ["Yeison Urrego", "Yeison Alonso Urrego Cortes"]
  },
  "dilma.cuesta@provexpress.com.co": {
    role: "comercial",
    comercial: "Dilma Cuesta",
    comercialAliases: ["Dilma Cuesta", "Dilma Constanza Cuesta Rubiano"],
    subRenta: true
  },
  "karen.carrillo@provexpress.com.co": {
    role: "comercial",
    comercial: "Karent Carrillo",
    comercialAliases: ["Karent Carrillo", "Karent Yessenia Carrillo Marin", "Karen Carrillo"],
    subRenta: true
  },
  "tatiana.parra@provexpress.com.co": {
    role: "comercial",
    comercial: "Tatiana Parra",
    comercialAliases: ["Tatiana Parra", "Angie Tatiana Parra Duran", "Angie Parra"]
  },
  "johanna.jaime@provexpress.com.co": {
    role: "comercial",
    comercial: "Johanna Jaime",
    comercialAliases: ["Johanna Jaime", "Johanna Jaime Murcia"]
  },
  "javier.cortes@provexpress.com.co": {
    role: "comercial",
    comercial: "Javier Cortes",
    comercialAliases: ["Javier Cortes", "Javier Antonio Cortes Murcia"]
  },
  "daniel.galindo@provexpress.com.co": {
    role: "comercial",
    comercial: "Daniel Galindo",
    comercialAliases: ["Daniel Galindo", "Daniel Galindo Giron"]
  },
  "julieth.galindo@provexpress.com.co": {
    role: "comercial",
    comercial: "Julieth Galindo",
    comercialAliases: ["Julieth Galindo", "Julieth Milena Galindo Fino"]
  },
  "claudia.triana@provexpress.com.co": {
    role: "comercial",
    comercial: "Claudia Triana",
    comercialAliases: ["Claudia Triana", "Claudia Patricia Triana Olaya"]
  },
  "fernando.quinonez@provexpress.com.co": {
    role: "comercial",
    comercial: "Fernando Quinonez",
    comercialAliases: ["Fernando Quinonez", "Fernando Alberto Quinonez"]
  },
  "camilo.hernandez@provexpress.com.co": {
    role: "comercial",
    comercial: "Camilo Hernandez",
    comercialAliases: ["Camilo Hernandez", "Jhonatan Camilo Hernandez Martinez", "Jhonatan Hernandez"]
  },
  "diana.castro@provexpress.com.co": {
    role: "comercial",
    comercial: "Diana Castro",
    comercialAliases: ["Diana Castro", "Diana Catalina Castro Castro"]
  },
  "juan.velasquez@provexpress.com.co": {
    role: "comercial",
    comercial: "Juan Velasquez",
    comercialAliases: ["Juan Velasquez", "Juan Camilo Velasquez Graciano"]
  },
  "andres.pena@provexpress.com.co": {
    role: "comercial",
    comercial: "Andres Pena",
    comercialAliases: ["Andres Pena", "Freddy Andres Pena Sanchez", "Freddy Pena"]
  }
};

(function () {
  function getCurrentUserAccess(accountOrProfile) {
    const email = getEmail(accountOrProfile);
    const access = USER_ACCESS[email];

    if (!access) {
      return {
        role: "sin_permiso",
        comercial: null,
        email
      };
    }

    return {
      ...access,
      email
    };
  }

  function getEmail(accountOrProfile) {
    return (accountOrProfile.mail || accountOrProfile.userPrincipalName || accountOrProfile.username || "").toLowerCase();
  }

  function isAdmin(user) {
    return ["gerencia", "finanzas", "operaciones"].includes(user.role);
  }

  function isCommercial(user) {
    return user.role === "comercial";
  }

  function canViewFinancials(user) {
    return ["gerencia", "finanzas", "comercial"].includes(user.role);
  }

  function canViewCommercialRanking(user) {
    return ["gerencia", "finanzas"].includes(user.role);
  }

  function canViewGlobalDashboard(user) {
    return ["gerencia", "finanzas", "operaciones"].includes(user.role);
  }

  function canViewSubRent(user) {
    return user.role === "gerencia" || Boolean(user.subRenta);
  }

  function getScopedRows(user, rows) {
    if (user.role === "comercial") {
      const aliases = user.comercialAliases || [user.comercial];
      const allowedNames = aliases.map((name) => ExcelService.comparableText(name));
      return rows.filter((row) => allowedNames.includes(ExcelService.comparableText(row.comercial)));
    }

    return rows;
  }

  window.PermissionService = {
    getCurrentUserAccess,
    isAdmin,
    isCommercial,
    canViewFinancials,
    canViewCommercialRanking,
    canViewGlobalDashboard,
    canViewSubRent,
    getScopedRows
  };
})();
