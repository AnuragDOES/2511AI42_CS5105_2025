# BTP/MTP Faculty Allocation System

A production-ready Streamlit application for automated student-faculty allocation using a fair mod-n algorithm.

## Features

✅ **Dynamic Faculty Detection** - Automatically counts faculty columns after CGPA  
✅ **CGPA-Based Sorting** - Sorts students in descending order of merit  
✅ **Mod n Allocation** - Round-robin distribution ensuring fair allocation  
✅ **Preference Statistics** - Comprehensive faculty preference analysis  
✅ **Error Handling** - Try-catch blocks with detailed logging  
✅ **Docker Support** - Complete containerization for easy deployment  
✅ **File Downloads** - Export allocation and statistics as CSV files  

## Quick Start

### Run Locally (Without Docker)

```bash
# Install dependencies
pip install -r requirements.txt

# Run the application
streamlit run app.py
```

Access at: http://localhost:8501

### Run with Docker (Recommended)

```bash
# Build and start
docker-compose up --build

# Or run in background
docker-compose up -d
```

Access at: http://localhost:8501

## Usage

1. **Upload CSV File** with format:
   - Columns: `Roll, Name, Email, CGPA, FAC1, FAC2, ...`
   - Faculty columns contain preference rankings (1 = top choice)

2. **Click "Process Allocation"**

3. **Download Results**:
   - `output_btp_mtp_allocation.csv` - Student assignments
   - `fac_preference_count.csv` - Preference statistics

## Algorithm

1. **Sort** students by CGPA (descending)
2. **Allocate** using mod n: Student at position i gets faculty at index (i mod n)
3. **Ensure** fair distribution - each faculty gets equal students per cycle

## File Structure

```
.
├── app.py                 # Main Streamlit application
├── requirements.txt       # Python dependencies
├── Dockerfile            # Docker configuration
├── docker-compose.yml    # Docker Compose setup
├── README.md             # This file
└── app.log              # Application logs (generated)
```

## Troubleshooting

**Port already in use:**
```bash
# Change port in docker-compose.yml or:
docker-compose down
```

**Module not found:**
```bash
pip install -r requirements.txt
```

**Check logs:**
```bash
# Docker logs
docker-compose logs -f

# Local logs
cat app.log
```

## Technical Details

- **Language**: Python 3.11
- **Framework**: Streamlit
- **Data Processing**: Pandas
- **Time Complexity**: O(n log n) - dominated by sorting
- **Space Complexity**: O(n × m) where n=students, m=faculties

## Requirements Met

✅ Dynamic faculty count (mod n)  
✅ Runtime CGPA sorting  
✅ Two CSV outputs (allocation + statistics)  
✅ Streamlit file upload/download  
✅ Docker + docker-compose  
✅ Works with/without Docker  
✅ Python logging library  
✅ Try-catch error handling  

## License

MIT License - Free to use and modify

## Support

For issues or questions, check the application logs in `app.log`
