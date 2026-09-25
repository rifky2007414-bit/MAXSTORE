const SPREADSHEET_ID=SpreadsheetApp.getActiveSpreadsheet().getId();
const COMMISSION_RATE=0.10;
function out(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON)}
function sh(n){return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(n)}
function rows(n){const s=sh(n);if(!s||s.getLastRow()<2)return[];const v=s.getDataRange().getValues(),h=v.shift();return v.map(r=>Object.fromEntries(h.map((k,i)=>[String(k),r[i]])))}
function uid(p){return p+"_"+Date.now()+"_"+Math.floor(Math.random()*10000)}
function hash(v){return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(v),Utilities.Charset.UTF_8).map(b=>("0"+(b&255).toString(16)).slice(-2)).join("")}
function doGet(e){try{const a=e.parameter.action,p=e.parameter;
if(a==="products")return out({success:true,products:rows("Products")});
if(a==="categories")return out({success:true,categories:rows("Categories")});
if(a==="sellers")return out({success:true,sellers:rows("Sellers")});
if(a==="users")return out({success:true,users:rows("Users")});
if(a==="orders"){let x=rows("Orders");if(p.userId)x=x.filter(r=>String(r.UserID)===String(p.userId));if(p.sellerId)x=x.filter(r=>String(r.SellerID)===String(p.sellerId));return out({success:true,orders:x})}
if(a==="cart")return out({success:true,cart:rows("Cart").filter(r=>String(r.UserID)===String(p.userId))});
if(a==="wishlist")return out({success:true,wishlist:rows("Wishlist").filter(r=>String(r.UserID)===String(p.userId))});
if(a==="reviews")return out({success:true,reviews:rows("Reviews")});
if(a==="notifications")return out({success:true,notifications:rows("Notifications").filter(r=>!p.userId||String(r.UserID)===String(p.userId))});
if(a==="withdrawals")return out({success:true,withdrawals:rows("Withdrawals").filter(r=>!p.sellerId||String(r.SellerID)===String(p.sellerId))});
if(a==="sellerBalances"){const sid=p.sellerId;const sales=rows("Orders").filter(r=>String(r.SellerID)===String(sid)&&String(r.OrderStatus).toLowerCase()==="delivered").reduce((t,r)=>t+Number(r.TotalAmount||0),0);const approved=rows("Withdrawals").filter(r=>String(r.SellerID)===String(sid)&&String(r.Status).toLowerCase()==="approved").reduce((t,r)=>t+Number(r.Amount||0),0);return out({success:true,balance:{AvailableBalance:Math.max(0,sales*(1-COMMISSION_RATE)-approved),TotalSales:sales,TotalCommission:sales*COMMISSION_RATE}})}
return out({success:false,message:"Unknown action"})}catch(err){return out({success:false,message:String(err)})}}
function doPost(e){let d;try{d=JSON.parse(e.postData.contents||"{}")}catch(err){return out({success:false,message:"Invalid JSON"})}try{const a=d.action;
if(a==="register"){const s=sh("Users"),u=rows("Users");if(u.some(x=>String(x.Email).toLowerCase()===String(d.email).toLowerCase()))return out({success:false,message:"Email already registered"});s.appendRow([uid("USR"),d.name,d.email,d.phone,hash(d.password)]);return out({success:true,message:"Account created successfully!"})}
if(a==="login"){const u=rows("Users").find(x=>String(x.Email).toLowerCase()===String(d.email).toLowerCase()&&String(x.PasswordHash)===hash(d.password));if(!u)return out({success:false,message:"Invalid email or password"});const id=u["User ID"]||u.UserID,s=rows("Sellers").find(x=>String(x.UserID)===String(id));const role=String(u.Email).toLowerCase()==="admin@maxstore.lk"?"admin":s?"seller":"buyer";return out({success:true,user:{userId:id,name:u.Name,email:u.Email,role,sellerId:s?.SellerID||""}})}
if(a==="addProduct"){sh("Products").appendRow([d.productID||uid("PRD"),d.ProductName||"",d.MainCategory||"",d.SubCategory||"",Number(d.Price||0),Number(d.Stock||0),d.sellerID||"",d.ImageURL||"",d.Description||"",d.Status||"Active"]);return out({success:true,message:"Product added"})}
if(a==="updateProduct"){const s=sh("Products"),v=s.getDataRange().getValues(),h=v[0],ix=h.indexOf("ProductID");for(let i=1;i<v.length;i++)if(String(v[i][ix])===String(d.productID)){["ProductName","MainCategory","Price","Stock","ImageURL","Description","Status"].forEach(k=>{const j=h.indexOf(k);if(j>=0&&d[k]!==undefined)s.getRange(i+1,j+1).setValue(d[k])});return out({success:true,message:"Product updated"})}return out({success:false,message:"Product not found"})}
if(a==="deleteProduct"){const s=sh("Products"),v=s.getDataRange().getValues(),h=v[0],ix=h.indexOf("ProductID");for(let i=v.length-1;i>0;i--)if(String(v[i][ix])===String(d.productID)){s.deleteRow(i+1);return out({success:true,message:"Product deleted"})}return out({success:false,message:"Product not found"})}
if(a==="addCart"){sh("Cart").appendRow([uid("CRT"),d.userId,d.productId,Number(d.quantity||1),new Date()]);return out({success:true,message:"Added to cart"})}
if(a==="withdraw"){sh("Withdrawals").appendRow([uid("WD"),d.sellerID||d.sellerId,Number(d.amount||0),d.method||"",d.details||"","Pending",new Date()]);return out({success:true,message:"Withdrawal request submitted"})}
if(a==="updateOrderStatus"){const s=sh("Orders"),v=s.getDataRange().getValues(),h=v[0],ix=h.indexOf("OrderID"),st=h.indexOf("OrderStatus");for(let i=1;i<v.length;i++)if(String(v[i][ix])===String(d.orderId)){s.getRange(i+1,st+1).setValue(d.status);return out({success:true,message:"Order updated"})}return out({success:false,message:"Order not found"})}
return out({success:false,message:"Unknown action"})}catch(err){return out({success:false,message:String(err)})}}
