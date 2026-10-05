(function () {
  "use strict";

  var box = document.getElementById("wrenchoo-ratings");
  if (!box) return;

  var slug = box.getAttribute("data-tool-slug");
  if (!slug) return;

  var stars = Array.prototype.slice.call(box.querySelectorAll(".wr-rating-star"));
  var review = box.querySelector(".wr-rating-review");
  var writeButton = box.querySelector(".wr-rating-write");
  var form = box.querySelector(".wr-rating-form");
  var submit = box.querySelector(".wr-rating-submit");
  var cancel = box.querySelector(".wr-rating-cancel");
  var message = box.querySelector(".wr-rating-message");
  var summary = box.querySelector(".wr-ratings-summary");
  var reviewsBox = box.querySelector(".wr-reviews");
  var selected = 0;
  var savedRating = 0;
  var busy = false;

  function visitorId() {
    var key = "wrenchoo_visitor_id";
    try {
      var value = localStorage.getItem(key);
      if (!value) {
        value = (window.crypto && crypto.randomUUID)
          ? crypto.randomUUID().replace(/-/g, "")
          : (Date.now().toString(36) + Math.random().toString(36).slice(2));
        localStorage.setItem(key, value);
      }
      return value;
    } catch (e) {
      return (Date.now().toString(36) + Math.random().toString(36).slice(2));
    }
  }

  var id = visitorId();

  function drawStars(value) {
    stars.forEach(function (star) {
      var n = Number(star.getAttribute("data-rating"));
      star.classList.toggle("is-selected", n <= value);
      star.classList.toggle("is-preview", n <= value);
      star.setAttribute("aria-pressed", n === value ? "true" : "false");
    });
  }

  function setMessage(text, type) {
    message.textContent = text || "";
    message.className = "wr-rating-message" + (type ? " is-" + type : "");
  }

  function render(data) {
    var count = Number(data.count || 0);
    var avg = Number(data.average || 0);
    summary.textContent = count
      ? (avg.toFixed(1) + " ★ · " + count + (count === 1 ? " rating" : " ratings"))
      : "";

    reviewsBox.innerHTML = "";
    if (!data.reviews || !data.reviews.length) return;

    var title = document.createElement("h3");
    title.className = "wr-reviews-title";
    title.textContent = "Recent reviews";
    reviewsBox.appendChild(title);

    data.reviews.forEach(function (item) {
      var row = document.createElement("div");
      row.className = "wr-review";

      var top = document.createElement("div");
      top.className = "wr-review-top";

      var rating = document.createElement("span");
      rating.className = "wr-review-stars";
      rating.textContent = "★".repeat(Math.max(1, Math.min(5, Number(item.rating) || 0)));

      var date = document.createElement("span");
      date.className = "wr-review-date";
      date.textContent = item.date || "";

      top.appendChild(rating);
      top.appendChild(date);

      var text = document.createElement("p");
      text.className = "wr-review-text";
      text.textContent = item.review;

      row.appendChild(top);
      if (item.review) row.appendChild(text);
      reviewsBox.appendChild(row);
    });
  }

  function saveRating(reviewText, successMessage) {
    if (!selected || busy) return Promise.reject(new Error("Please choose a star rating first."));
    busy = true;
    setMessage("", "");

    return fetch("/api/ratings/" + encodeURIComponent(slug), {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        visitor_id: id,
        rating: selected,
        review: (reviewText || "").slice(0, 500)
      }),
      cache: "no-store"
    })
      .then(function (response) {
        return response.json().then(function (data) {
          if (!response.ok || !data.ok) throw new Error(data.error || "Could not save feedback.");
          return data;
        });
      })
      .then(function (data) {
        savedRating = selected;
        render(data);
        if (successMessage) setMessage(successMessage, "success");
      })
      .catch(function (error) {
        setMessage(error.message || "Your feedback could not be saved right now.", "error");
      })
      .finally(function () {
        busy = false;
      });
  }

  stars.forEach(function (star) {
    star.addEventListener("mouseenter", function () {
      drawStars(Number(star.getAttribute("data-rating")));
    });
    star.addEventListener("focus", function () {
      drawStars(Number(star.getAttribute("data-rating")));
    });
    star.addEventListener("mouseleave", function () {
      drawStars(selected);
    });
    star.addEventListener("blur", function () {
      drawStars(selected);
    });
    star.addEventListener("click", function () {
      selected = Number(star.getAttribute("data-rating"));
      drawStars(selected);
      saveRating("", "Thanks for rating this tool!");
    });
  });

  writeButton.addEventListener("click", function () {
    if (!selected) {
      selected = savedRating || 0;
      if (selected) drawStars(selected);
    }
    form.classList.add("is-open");
    review.focus();
    writeButton.setAttribute("aria-expanded", "true");
  });

  cancel.addEventListener("click", function () {
    form.classList.remove("is-open");
    review.value = "";
    setMessage("", "");
    writeButton.setAttribute("aria-expanded", "false");
  });

  submit.addEventListener("click", function () {
    if (!selected) {
      setMessage("Please choose a star rating first.", "error");
      return;
    }
    submit.disabled = true;
    saveRating(review.value, "Thanks for your review!")
      .then(function () {
        review.value = "";
        form.classList.remove("is-open");
        writeButton.setAttribute("aria-expanded", "false");
      })
      .catch(function () {})
      .finally(function () {
        submit.disabled = false;
      });
  });

  fetch("/api/ratings/" + encodeURIComponent(slug), { cache: "no-store" })
    .then(function (response) { return response.ok ? response.json() : null; })
    .then(function (data) { if (data && data.ok) render(data); })
    .catch(function () { /* Ratings are optional; don't affect the tool. */ });
})();
