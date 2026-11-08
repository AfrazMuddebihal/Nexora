from flask import Flask,render_template,request,redirect,url_for,send_file,flash
import sqlite3,io,os,json,urllib.request,urllib.error
from datetime import date,datetime
from openpyxl import Workbook
app=Flask(__name__); app.secret_key=os.environ.get("SECRET_KEY","nexora-demo"); DB=os.environ.get("DATABASE_PATH","nexora.db")
def conn():
 c=sqlite3.connect(DB); c.row_factory=sqlite3.Row; return c
def init():
 with conn() as c:
  c.execute("CREATE TABLE IF NOT EXISTS students(id INTEGER PRIMARY KEY,roll TEXT UNIQUE,name TEXT,active INTEGER DEFAULT 1)")
  c.execute("CREATE TABLE IF NOT EXISTS attendance(student_id INTEGER,day TEXT,status TEXT,UNIQUE(student_id,day))")
  # Safe migration for older NEXORA databases
  cols=[r[1] for r in c.execute("PRAGMA table_info(students)")]
  if 'parent_phone' not in cols:c.execute("ALTER TABLE students ADD COLUMN parent_phone TEXT")
  c.execute("CREATE TABLE IF NOT EXISTS sms_log(id INTEGER PRIMARY KEY,student_id INTEGER,day TEXT,status TEXT,detail TEXT,created_at TEXT,UNIQUE(student_id,day))")
init()
@app.route("/")
def home():
 d=request.args.get("day",date.today().isoformat())
 with conn() as c:
  total=c.execute("SELECT count(*) FROM students WHERE active=1").fetchone()[0]
  p=c.execute("SELECT count(*) FROM attendance a JOIN students s ON s.id=a.student_id WHERE day=? AND status='Present' AND active=1",(d,)).fetchone()[0]
  ab=c.execute("SELECT count(*) FROM attendance a JOIN students s ON s.id=a.student_id WHERE day=? AND status='Absent' AND active=1",(d,)).fetchone()[0]
  finalized=c.execute("SELECT count(*) FROM sms_log WHERE day=?",(d,)).fetchone()[0]>0
 return render_template("home.html",day=d,total=total,p=p,ab=ab,left=max(0,total-p-ab),finalized=finalized)
@app.route("/students",methods=["GET","POST"])
def students():
 with conn() as c:
  if request.method=="POST":
   try:c.execute("INSERT INTO students(roll,name,parent_phone) VALUES(?,?,?)",(request.form["roll"].strip(),request.form["name"].strip(),request.form.get("parent_phone"," ").strip()))
   except sqlite3.IntegrityError:flash("Roll number already exists")
   return redirect(url_for("students"))
  rows=c.execute("SELECT * FROM students ORDER BY roll").fetchall()
 return render_template("students.html",rows=rows)
@app.route("/register")
def register():
 d=request.args.get("day",date.today().isoformat())
 with conn() as c: rows=c.execute("SELECT s.*,a.status FROM students s LEFT JOIN attendance a ON a.student_id=s.id AND a.day=? WHERE s.active=1 ORDER BY roll",(d,)).fetchall(); finalized=c.execute("SELECT count(*) FROM sms_log WHERE day=?",(d,)).fetchone()[0]>0
 return render_template("register.html",rows=rows,day=d,finalized=finalized)
@app.route("/mark/<int:i>/<status>",methods=["POST"])
def mark(i,status):
 if status not in ('Present','Absent'):return 'Invalid status',400
 d=request.form.get("day",date.today().isoformat())
 with conn() as c:
  if c.execute("SELECT count(*) FROM sms_log WHERE day=?",(d,)).fetchone()[0]:flash("Attendance already finalized for this date. Reopen it before editing.");return redirect(url_for("register",day=d))
  c.execute("INSERT INTO attendance VALUES(?,?,?) ON CONFLICT(student_id,day) DO UPDATE SET status=excluded.status",(i,d,status))
 return redirect(url_for("register",day=d))
def send_sms(phone,name,day):
 auth=os.environ.get('MSG91_AUTHKEY'); flow=os.environ.get('MSG91_FLOW_ID'); sender=os.environ.get('MSG91_SENDER')
 if not (auth and flow and sender):return 'DEMO','Provider credentials not configured; no SMS sent'
 payload={'flow_id':flow,'sender':sender,'mobiles':phone,'student':name,'date':day}
 req=urllib.request.Request('https://api.msg91.com/api/v5/flow/',data=json.dumps(payload).encode(),headers={'authkey':auth,'Content-Type':'application/json'},method='POST')
 try:
  with urllib.request.urlopen(req,timeout=15) as r:return ('SENT' if r.status in (200,201) else 'FAILED'),r.read().decode()[:500]
 except Exception as e:return 'FAILED',str(e)[:500]
@app.route('/finalize/<day>',methods=['POST'])
def finalize(day):
 with conn() as c:
  if c.execute('SELECT count(*) FROM sms_log WHERE day=?',(day,)).fetchone()[0]:flash('This date is already finalized.');return redirect(url_for('register',day=day))
  absent=c.execute("SELECT s.id,s.name,s.parent_phone FROM students s JOIN attendance a ON a.student_id=s.id WHERE s.active=1 AND a.day=? AND a.status='Absent'",(day,)).fetchall()
  marked=c.execute("SELECT count(*) FROM attendance a JOIN students s ON s.id=a.student_id WHERE s.active=1 AND a.day=?",(day,)).fetchone()[0]
  total=c.execute('SELECT count(*) FROM students WHERE active=1').fetchone()[0]
  if marked<total:flash('Please mark every active student before finalizing.');return redirect(url_for('register',day=day))
 for s in absent:
  phone=(s['parent_phone'] or '').strip()
  if not phone:status,detail='NO NUMBER','Parent mobile number not entered'
  else:status,detail=send_sms(phone,s['name'],day)
  with conn() as c:c.execute('INSERT OR REPLACE INTO sms_log(student_id,day,status,detail,created_at) VALUES(?,?,?,?,?)',(s['id'],day,status,detail,datetime.now().isoformat(timespec='seconds')))
 flash(f'Attendance finalized. Parent alerts processed for {len(absent)} absent student(s). Check SMS status below.')
 return redirect(url_for('register',day=day))
@app.route('/reopen/<day>',methods=['POST'])
def reopen(day):
 with conn() as c:c.execute('DELETE FROM sms_log WHERE day=?',(day,))
 flash('Date reopened. You can edit attendance and finalize again.');return redirect(url_for('register',day=day))
@app.route('/export')
def export():
 d=request.args.get('day',date.today().isoformat())
 with conn() as c: rows=c.execute("SELECT roll,name,COALESCE(status,'Not marked') FROM students LEFT JOIN attendance ON id=student_id AND day=? WHERE active=1 ORDER BY roll",(d,)).fetchall()
 w=Workbook();s=w.active;s.append(["Roll","Name","Status"]);[s.append(list(x)) for x in rows]
 b=io.BytesIO();w.save(b);b.seek(0);return send_file(b,as_attachment=True,download_name=f"NEXORA-{d}.xlsx")
@app.route('/sms-status')
def sms_status():
 d=request.args.get('day',date.today().isoformat())
 with conn() as c:rows=c.execute('SELECT s.roll,s.name,s.parent_phone,l.status,l.detail,l.created_at FROM sms_log l JOIN students s ON s.id=l.student_id WHERE l.day=? ORDER BY s.roll',(d,)).fetchall()
 return render_template('sms_status.html',rows=rows,day=d)
if __name__=='__main__':app.run(host='0.0.0.0',port=int(os.environ.get('PORT',5000)))
