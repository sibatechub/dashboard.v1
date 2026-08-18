/*************************************************
 * FEEDBACK BACKEND
 * Source Sheet: Master
 *************************************************/


/*************************************************
 * GET FEEDBACK DATA
 *************************************************/

function getFeedbackData(sessionToken, filters) {

  const user = getUserSession(sessionToken);

  if (!user) {
    throw new Error("Session expired. Please login again.");
  }

  const role = String(user.role || "")
      .trim()
      .toUpperCase();

  const userName = String(user.name || "")
      .trim()
      .toUpperCase();


  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  const sheet = ss.getSheetByName("Master");

  if (!sheet) {
    throw new Error("Master sheet not found.");
  }


  const data = sheet.getDataRange().getValues();

  if (data.length <= 1) {

    return {

      records: [],
      summary: {

        totalFeedback: 0,
        averageRating: 0,

        behaviour: "-",
        behaviourCount: 0,

        resolution: "-",
        resolutionCount: 0

      },

      technicians: []

    };

  }


  data.shift();


  filters = filters || {};

  const search =
      String(filters.search || "")
      .trim()
      .toLowerCase();

  const buildingFilter =
      String(filters.building || "")
      .trim()
      .toLowerCase();

  const technicianFilter =
      String(filters.technician || "")
      .trim()
      .toLowerCase();

  const ratingFilter =
      String(filters.rating || "")
      .trim();


  const records = [];


  /*************************************************
   * PROCESS MASTER DATA
   *************************************************/

  data.forEach(function(row) {


    /*
     * Overall Satisfaction
     * Column W = index 22
     */

    const ratingValue = row[22];


    /*
     * Ignore rows where feedback has not
     * been submitted.
     */

    if (
      ratingValue === "" ||
      ratingValue === null ||
      ratingValue === undefined
    ) {

      return;

    }


    const rating = Number(ratingValue);


    if (isNaN(rating)) {
      return;
    }


    /*
     * Get values
     */

    const ticketNo =
        String(row[1] || "");

    const faculty =
        String(row[5] || "");

    const building =
        String(row[8] || "");

    const technician =
        String(row[15] || "");

    const behaviour =
        String(row[23] || "");

    const resolution =
        String(row[24] || "");

    const comments =
        String(row[25] || "");

    const feedbackDate =
        row[30];


    /*************************************************
     * TECHNICIAN LOGIN SECURITY
     *************************************************/

    if (role !== "ADMIN") {

      if (
        technician.trim().toUpperCase() !== userName
      ) {

        return;

      }

    }


    /*************************************************
     * BUILDING FILTER
     *************************************************/

    if (
      buildingFilter &&
      building.toLowerCase() !== buildingFilter
    ) {

      return;

    }


    /*************************************************
     * TECHNICIAN FILTER
     *************************************************/

    if (
      technicianFilter &&
      technician.toLowerCase() !== technicianFilter
    ) {

      return;

    }


    /*************************************************
     * RATING FILTER
     *************************************************/

    if (
      ratingFilter &&
      String(rating) !== ratingFilter
    ) {

      return;

    }


    /*************************************************
     * SEARCH
     *************************************************/

    if (search) {

      const searchableText = (

        ticketNo + " " +
        faculty + " " +
        building + " " +
        technician + " " +
        behaviour + " " +
        resolution + " " +
        comments

      ).toLowerCase();


      if (
        !searchableText.includes(search)
      ) {

        return;

      }

    }


    /*************************************************
     * FORMAT DATE
     *************************************************/

    let formattedDate = "-";

    if (feedbackDate) {

      try {

        formattedDate =
            Utilities.formatDate(
                new Date(feedbackDate),
                Session.getScriptTimeZone(),
                "dd/MM/yyyy hh:mm a"
            );

      } catch (e) {

        formattedDate =
            String(feedbackDate);

      }

    }


    /*************************************************
     * ADD RECORD
     *************************************************/

    records.push({

      ticketNo: ticketNo,

      faculty: faculty,

      building: building,

      technician: technician,

      rating: rating,

      behaviour: behaviour || "-",

      resolution: resolution || "-",

      comments: comments || "-",

      feedbackDate: formattedDate

    });

  });


  /*************************************************
   * SORT
   * Latest feedback first
   *************************************************/

  records.reverse();


  /*************************************************
   * SUMMARY
   *************************************************/

  const summary =
      calculateFeedbackSummary(records);


  /*************************************************
   * TECHNICIAN PERFORMANCE
   *************************************************/

  const technicians =
      calculateTechnicianPerformance(records);


  return {

    records: records,

    summary: summary,

    technicians: technicians

  };

}


/*************************************************
 * FEEDBACK SUMMARY
 *************************************************/

function calculateFeedbackSummary(records) {


  const totalFeedback =
      records.length;


  if (totalFeedback === 0) {

    return {

      totalFeedback: 0,

      averageRating: 0,

      behaviour: "-",
      behaviourCount: 0,

      resolution: "-",
      resolutionCount: 0

    };

  }


  /*************************************************
   * AVERAGE RATING
   *************************************************/

  let totalRating = 0;


  records.forEach(function(record) {

    totalRating +=
        Number(record.rating) || 0;

  });


  const averageRating =
      totalRating / totalFeedback;


  /*************************************************
   * MOST COMMON BEHAVIOUR
   *************************************************/

  const behaviourResult =
      getMostCommonValue(
          records,
          "behaviour"
      );


  /*************************************************
   * MOST COMMON RESOLUTION
   *************************************************/

  const resolutionResult =
      getMostCommonValue(
          records,
          "resolution"
      );


  return {

    totalFeedback: totalFeedback,

    averageRating:
        Number(averageRating.toFixed(2)),

    behaviour:
        behaviourResult.value,

    behaviourCount:
        behaviourResult.count,

    resolution:
        resolutionResult.value,

    resolutionCount:
        resolutionResult.count

  };

}


/*************************************************
 * MOST COMMON TEXT VALUE
 *************************************************/

function getMostCommonValue(records, property) {


  const countMap = {};


  records.forEach(function(record) {


    const value =
        String(record[property] || "")
        .trim();


    if (
      !value ||
      value === "-"
    ) {

      return;

    }


    const key =
        value.toLowerCase();


    if (!countMap[key]) {

      countMap[key] = {

        value: value,

        count: 0

      };

    }


    countMap[key].count++;

  });


  let result = {

    value: "-",

    count: 0

  };


  Object.keys(countMap).forEach(function(key) {


    if (
      countMap[key].count >
      result.count
    ) {

      result = {

        value: countMap[key].value,

        count: countMap[key].count

      };

    }

  });


  return result;

}


/*************************************************
 * TECHNICIAN PERFORMANCE
 *************************************************/

function calculateTechnicianPerformance(records) {


  const technicianMap = {};


  records.forEach(function(record) {


    const name =
        String(record.technician || "")
        .trim();


    if (!name) {
      return;
    }


    const key =
        name.toLowerCase();


    if (!technicianMap[key]) {

      technicianMap[key] = {

        name: name,

        feedbacks: 0,

        totalRating: 0,

        behaviours: {},

        resolutions: {}

      };

    }


    const tech =
        technicianMap[key];


    tech.feedbacks++;


    tech.totalRating +=
        Number(record.rating) || 0;


    /*************************************************
     * BEHAVIOUR COUNT
     *************************************************/

    const behaviour =
        String(record.behaviour || "")
        .trim();


    if (behaviour && behaviour !== "-") {

      const behaviourKey =
          behaviour.toLowerCase();


      if (!tech.behaviours[behaviourKey]) {

        tech.behaviours[behaviourKey] = {

          value: behaviour,

          count: 0

        };

      }


      tech.behaviours[behaviourKey].count++;

    }


    /*************************************************
     * RESOLUTION COUNT
     *************************************************/

    const resolution =
        String(record.resolution || "")
        .trim();


    if (resolution && resolution !== "-") {

      const resolutionKey =
          resolution.toLowerCase();


      if (!tech.resolutions[resolutionKey]) {

        tech.resolutions[resolutionKey] = {

          value: resolution,

          count: 0

        };

      }


      tech.resolutions[resolutionKey].count++;

    }

  });


  const result = [];


  Object.keys(technicianMap).forEach(function(key) {


    const tech =
        technicianMap[key];


    const averageRating =
        tech.feedbacks > 0

        ? tech.totalRating / tech.feedbacks

        : 0;


    const behaviour =
        getMostCommonObject(
            tech.behaviours
        );


    const resolution =
        getMostCommonObject(
            tech.resolutions
        );


    result.push({

      name: tech.name,

      feedbacks: tech.feedbacks,

      averageRating:
          Number(
              averageRating.toFixed(2)
          ),

      behaviour:
          behaviour.value,

      behaviourCount:
          behaviour.count,

      resolution:
          resolution.value,

      resolutionCount:
          resolution.count

    });

  });


  /*************************************************
   * HIGHEST RATING FIRST
   *************************************************/

  result.sort(function(a, b) {

    return b.averageRating -
           a.averageRating;

  });


  return result;

}


/*************************************************
 * MOST COMMON OBJECT
 *************************************************/

function getMostCommonObject(map) {


  let result = {

    value: "-",

    count: 0

  };


  Object.keys(map).forEach(function(key) {


    if (
      map[key].count >
      result.count
    ) {

      result = {

        value: map[key].value,

        count: map[key].count

      };

    }

  });


  return result;

}


/*************************************************
 * GET FILTER OPTIONS
 *************************************************/

function getFeedbackFilters(sessionToken) {


  const user =
      getUserSession(sessionToken);


  if (!user) {

    throw new Error(
        "Session expired. Please login again."
    );

  }


  const role =
      String(user.role || "")
      .trim()
      .toUpperCase();


  const userName =
      String(user.name || "")
      .trim()
      .toUpperCase();


  const ss =
      SpreadsheetApp.openById(
          SPREADSHEET_ID
      );


  const sheet =
      ss.getSheetByName("Master");


  const data =
      sheet.getDataRange()
      .getValues();


  data.shift();


  const buildings = {};
  const technicians = {};


  data.forEach(function(row) {


    const rating =
        row[22];


    /*
     * Only rows having feedback
     */

    if (
      rating === "" ||
      rating === null ||
      rating === undefined
    ) {

      return;

    }


    const building =
        String(row[8] || "")
        .trim();


    const technician =
        String(row[15] || "")
        .trim();


    /*************************************************
     * TECHNICIAN LOGIN
     *************************************************/

    if (role !== "ADMIN") {

      if (
        technician.toUpperCase() !==
        userName
      ) {

        return;

      }

    }


    if (building) {

      buildings[building] = true;

    }


    if (technician) {

      technicians[technician] = true;

    }

  });


  return {

    buildings:
        Object.keys(buildings).sort(),

    technicians:
        Object.keys(technicians).sort()

  };

}
