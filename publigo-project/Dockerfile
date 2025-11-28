# ใช้ Node.js v20 บน Linux
FROM node:20-slim

# 1. ติดตั้ง LibreOffice และฟอนต์ภาษาไทย (สำคัญมาก ไม่งั้นภาษาไทยจะเป็นสี่เหลี่ยม)
RUN apt-get update && apt-get install -y \
    libreoffice \
    fonts-thai-tlwg \
    fonts-google-license \
    && rm -rf /var/lib/apt/lists/*

# 2. ตั้งค่าโฟลเดอร์ทำงาน
WORKDIR /app

# 3. Copy ไฟล์สเปคและติดตั้งแพ็กเกจ
COPY package*.json ./
RUN npm install

# 4. Copy โค้ดทั้งหมดลงไป
COPY . .

# 5. เปิด Port 3000
EXPOSE 3000

# 6. สั่งรัน Server
CMD ["node", "server.js"]