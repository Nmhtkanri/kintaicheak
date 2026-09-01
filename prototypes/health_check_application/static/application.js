(() => {
  const form = document.querySelector('#application-form');
  if (!form) return;

  const changeFields = document.querySelector('#change-fields');
  const clinicChoice = form.elements.clinic_choice;
  const standardOptions = document.querySelector('#standard-options');
  const otherFields = document.querySelector('#other-fields');
  const optionChecks = [...form.querySelectorAll('input[name="health_options"]')];
  const dependentRequested = form.elements.dependent_requested;
  const dependentFields = document.querySelector('#dependent-fields');
  const relationship = form.elements.dependent_relationship;
  const dependentName = form.elements.dependent_name;
  const message = document.querySelector('#form-message');
  const submitButton = document.querySelector('#submit-button');
  const successPanel = document.querySelector('#success-panel');

  function selectedApplicationType() {
    const selected = form.querySelector('input[name="application_type"]:checked');
    return selected ? selected.value : '';
  }

  function setVisible(element, visible) {
    element.classList.toggle('is-hidden', !visible);
    element.setAttribute('aria-hidden', String(!visible));
  }

  function updateChangeFields() {
    const changing = selectedApplicationType() === 'change';
    setVisible(changeFields, changing);
    clinicChoice.required = changing;
    if (!changing) {
      clinicChoice.value = '';
      form.elements.custom_clinic.value = '';
      form.elements.other_planned_date.value = '';
      optionChecks.forEach((item) => { item.checked = false; });
    }
    updateClinicDetails();
  }

  function updateClinicDetails() {
    const changing = selectedApplicationType() === 'change';
    const choice = clinicChoice.value;
    const isOther = changing && choice === window.healthCheckPrototype.otherClinic;
    const isStandard = changing && Boolean(choice) && !isOther;
    setVisible(standardOptions, isStandard);
    setVisible(otherFields, isOther);
    if (!isStandard) optionChecks.forEach((item) => { item.checked = false; });
    if (!isOther) {
      form.elements.custom_clinic.value = '';
      form.elements.other_planned_date.value = '';
    }
    validateOptionChecks();
  }

  function validateOptionChecks() {
    if (!optionChecks.length) return;
    const requiresOption = !standardOptions.classList.contains('is-hidden');
    const hasOption = optionChecks.some((item) => item.checked);
    optionChecks[0].setCustomValidity(
      requiresOption && !hasOption
        ? '健診オプションを1つ以上選択してください。'
        : ''
    );
  }

  function updateDependentFields() {
    const visible = dependentRequested.checked;
    setVisible(dependentFields, visible);
    relationship.required = visible;
    dependentName.required = visible;
    if (!visible) {
      relationship.value = '';
      dependentName.value = '';
    }
  }

  form.querySelectorAll('input[name="application_type"]').forEach((radio) => {
    radio.addEventListener('change', updateChangeFields);
  });
  clinicChoice.addEventListener('change', updateClinicDetails);
  optionChecks.forEach((item) => item.addEventListener('change', validateOptionChecks));
  dependentRequested.addEventListener('change', updateDependentFields);

  form.addEventListener('submit', async (event) => {
    event.preventDefault();
    message.textContent = '';
    message.className = 'form-message';
    validateOptionChecks();
    if (!form.reportValidity()) return;

    const payload = {
      application_type: selectedApplicationType(),
      clinic_choice: clinicChoice.value,
      health_options: optionChecks.filter((item) => item.checked).map((item) => item.value),
      custom_clinic: form.elements.custom_clinic.value,
      other_planned_date: form.elements.other_planned_date.value,
      dependent_requested: dependentRequested.checked,
      dependent_relationship: relationship.value,
      dependent_name: dependentName.value,
      remarks: form.elements.remarks.value,
      agreement: form.elements.agreement.checked,
      // Identity is intentionally absent. The server derives it from the URL token.
    };

    submitButton.disabled = true;
    try {
      const response = await fetch(window.healthCheckPrototype.submitUrl, {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify(payload),
      });
      const result = await response.json();
      if (!response.ok) throw new Error(result.message || '送信できませんでした。');

      document.querySelector('#receipt-id').textContent = result.receipt_id;
      document.querySelector('#confirmed-employee-id').textContent = result.employee_id;
      form.classList.add('is-hidden');
      successPanel.classList.remove('is-hidden');
      successPanel.setAttribute('aria-hidden', 'false');
      successPanel.scrollIntoView({behavior: 'smooth', block: 'start'});
    } catch (error) {
      message.textContent = error.message;
      message.classList.add('form-message--error');
      submitButton.disabled = false;
    }
  });

  updateChangeFields();
  updateDependentFields();
})();

