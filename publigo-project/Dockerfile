# ใช้ Node.js v20 บน Linux
FROM node:20-slim

# 1. ติดตั้ง LibreOffice และฟอนต์ภาษาไทย
# (ลบ fonts-google-license ออก เพื่อแก้ Error)
RUN apt-get update && apt-get install -y \
    libreoffice \
    fonts-thai-tlwg \
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