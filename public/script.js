document.getElementById('dataForm').addEventListener('submit', async (event) => {
    event.preventDefault();
    const formData = {
        name: document.getElementById('name').value,
        age: document.getElementById('age').value,
    };
    const response = await fetch('/submit', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(formData),
    });
    alert(await response.text());
});
