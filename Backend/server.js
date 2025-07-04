const express=require('express');
const mongoose=require('mongoose');
const cors=require('cors');
const http=require('http');
const {Server}=require('socket.io')
require('dotenv').config();
const app=express();
const PORT=process.env.PORT||5000;
const MONGO_URI=process.env.MONGO_URI;
app.use(cors());
app.use(express.json({limit:'10mb'}));
app.use(express.urlencoded({extended:true,limit:'10mb'}));
mongoose.connect(MONGO_URI)
.then(()=>console.log('MongoDB connected successfully'))
.catch(err=>console.error('MongoDB connection error',err));

app.get('/',(req,res)=>{
    res.send('Spreadsheet Backend is running');
});
const sheetRoutes = require('./routes/Sheetroutes');
app.use('/api/sheets', sheetRoutes);
const server=http.createServer(app);
const io=new Server(server,{
    cors:{
    origin:"https://voluble-donut-c728cf.netlify.app",
    methods: ["GET", "POST"]
    },
})
io.on('connection',(socket)=>{
      console.log('User Connected')
      socket.on('cell-edit',(edit)=>{
        socket.broadcast.emit('cell-edit',edit)
      });
      socket.on('disconnect',()=>{
        console.log('User disconnected')
      })
})
server.listen(PORT,()=>{
    console.log(`Server running on port${PORT}`);
})