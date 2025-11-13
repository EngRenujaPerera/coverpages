function updatePreview() {
    document.getElementById("p_topic").innerText = document.getElementById("topic").value.toUpperCase();
    document.getElementById("p_name").innerText = document.getElementById("name").value.toUpperCase();
    document.getElementById("p_reg_no").innerText = document.getElementById("reg_no").value.toUpperCase();
    document.getElementById("p_instructor").innerText = document.getElementById("instructor").value.toUpperCase();
    document.getElementById("p_course_code").innerText = document.getElementById("course_code").value.toUpperCase();
    document.getElementById("p_exp_date").innerText = document.getElementById("exp_date").value.toUpperCase();
    document.getElementById("p_sub_date").innerText = document.getElementById("sub_date").value.toUpperCase();
}
