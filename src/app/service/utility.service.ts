import { Injectable } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { Observable } from 'rxjs';

@Injectable({
  providedIn: 'root'
})
export class UtilityService {

  constructor(private http: HttpClient) { }
  
  uploadDocToCommusoft(params: Object, data: Object): Observable<any>{
     return this.http.post('http://dev.v4.com/frontend_dev.php/upload_customer_file?mode=customers&selectedId=17', data)
  }
}
