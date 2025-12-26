# 1. ใช้ Node.js เวอร์ชันที่มี Linux พื้นฐานครบ (สำคัญมาก ห้ามใช้ Alpine)
FROM node:18-bullseye

# 2. ติดตั้ง LibreOffice และฟอนต์ภาษาไทย
RUN apt-get update && apt-get install -y \
    libreoffice \
    fonts-thai-tlwg \
    fonts-sarabun \
    && rm -rf /var/lib/apt/lists/*

# 3. กำหนดโฟลเดอร์ทำงาน
WORKDIR /app

# 4. ก๊อปปี้ไฟล์ package.json และติดตั้ง dependency
COPY package*.json ./
RUN npm install

# 5. ก๊อปปี้ไฟล์ทั้งหมดในโปรเจกต์
COPY . .

# 6. เปิด Port (Koyeb จะอ่าน Port จาก ENV หรือใช้ 8000)
ENV PORT=8000
EXPOSE 8000

# 7. คำสั่งรัน Server
CMD ["node", "server.js"]