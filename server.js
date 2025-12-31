// server.js
// 1) npm init -y
// 2) npm install express cors node-fetch
// 3) node server.js
// Open: http://localhost:4000

const express = require('express');
const cors = require('cors');
const path = require('path');
const fetch = require('node-fetch'); // will be used when you really call Graph

const app = express();
const PORT = 4000;

let tasks = [];
let completedTasks = [];

app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// ---------- Core task APIs ----------
app.get('/api/tasks', (req, res) => {
  tasks.sort((a, b) => new Date(a.deadline) - new Date(b.deadline));
  res.json(tasks);
});

app.get('/api/tasks/completed', (req, res) => {
  res.json(completedTasks);
});

app.post('/api/tasks', (req, res) => {
  const { title, type, deadline, subject, description } = req.body;
  if (!title || !type || !deadline) {
    return res.status(400).json({ error: 'title, type and deadline are required' });
  }

  const task = {
    id: Date.now().toString(),
    title,
    type,
    deadline,
    subject: subject || '',
    description: description || ''
  };

  tasks.push(task);
  tasks.sort((a, b) => new Date(a.deadline) - new Date(b.deadline));
  res.status(201).json(task);
});

app.post('/api/tasks/:id/complete', (req, res) => {
  const idx = tasks.findIndex(t => t.id === req.params.id);
  if (idx === -1) return res.status(404).json({ error: 'Task not found' });

  const task = tasks[idx];
  tasks.splice(idx, 1);
  completedTasks.push(task);
  res.json(task);
});

app.delete('/api/tasks/:id', (req, res) => {
  const idx = tasks.findIndex(t => t.id === req.params.id);
  if (idx === -1) return res.status(404).json({ error: 'Task not found' });

  const deleted = tasks.splice(idx, 1)[0];
  res.json(deleted);
});

// ---------- Microsoft To Do hook (placeholder) ----------
/*
 Real Graph call would look like:

 POST https://graph.microsoft.com/v1.0/me/todo/lists/{list-id}/tasks
 Authorization: Bearer ACCESS_TOKEN
 Body: { title: "...", body: { contentType: "text", content: "..." } }

 See docs: [web:138][web:141][web:219]
*/

app.post('/api/todo', async (req, res) => {
  const { title, type, deadline, subject, description } = req.body;

  // For now, just echo back what would be sent to Microsoft To Do
  // so you can see the payload shape.
  return res.json({
    status: 'demo',
    message: 'This is where Microsoft To Do Graph API would be called.',
    wouldSend: {
      title,
      body: {
        contentType: 'text',
        content: `Type: ${type}\nSubject: ${subject || ''}\nDeadline: ${deadline}\n\n${description || ''}`
      }
    }
  });
});

// ---------- Static frontend ----------
app.use(express.static(path.join(__dirname, 'public')));

app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.listen(PORT, () => {
  console.log(`DeadlineTracker app running at http://localhost:${PORT}`);
});
