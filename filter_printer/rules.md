ABSOLUTE RULES

- Do not rewrite the entire project.
- Do not reorganize folders.
- Do not split files.
- Do not rename public classes.
- Do not introduce new architecture.
- Do not optimize unrelated code.
- Keep the diff as small as possible.
- Every modification must have a clear reason.
- If a feature is already stable, leave it untouched.
- Backward compatibility is mandatory.

Bật chế độ Principal Engineer.

Không được giả định rằng code hiện tại là đúng.

Hãy cố gắng chứng minh code có vấn đề.

Với mỗi module, hãy tự hỏi:

- Có cách nào nhanh hơn 2 lần?
- Có cách nào giảm RAM 50%?
- Có cách nào giảm I/O?
- Có race condition tiềm ẩn?
- Có deadlock tiềm ẩn?
- Có memory leak?
- Có callback dư thừa?
- Có thread dư thừa?
- Có object tồn tại quá lâu?
- Có cache sai vị trí?
- Có thao tác O(N²) nào có thể thành O(N)?
- Có thao tác O(N) nào có thể thành O(logN)?
- Có thao tác nào nên lazy?
- Có thao tác nào nên streaming?
- Có thao tác nào nên bất đồng bộ?
- Có thao tác nào đang block UI?
- Có thể loại bỏ hoàn toàn thao tác đó không?

Không chỉ sửa theo yêu cầu.

Hãy chủ động tìm lỗi mà người viết code chưa nhận ra.

Ưu tiên các thay đổi mang lại lợi ích lớn nhất theo nguyên tắc Pareto (20% thay đổi tạo ra 80% cải thiện).

Chỉ dừng khi không còn tìm thấy bất kỳ cải tiến nào có ý nghĩa về hiệu năng, độ ổn định, khả năng mở rộng hoặc trải nghiệm người dùng.